#!/usr/bin/env python3
"""Fetch every common trading-day close through an as-of date."""

import argparse
import csv
import json
import re
from datetime import date, datetime, timedelta
from pathlib import Path

import yfinance as yf


DEFAULT_TICKERS = [
    "MSFT", "NVDA", "TSLA", "VOO", "RGTI", "IONQ",
    "ARM", "MSTR", "SATS", "AMZN", "GOOG",
]

# Keep workbook symbols stable when a company changes its market ticker.
# Yahoo serves EchoStar's price history under ECHO after the 2026-06-24 rename,
# while the existing workbook still identifies the holding as SATS.
YAHOO_SYMBOLS = {"SATS": "ECHO"}


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("--as-of", help="Requested date in YYYY-MM-DD format")
    parser.add_argument("--start", help="First requested date in YYYY-MM-DD format; defaults to the as-of month start")
    parser.add_argument("--cashflow-csv", action="append", help="Use the latest 日期 across one or more 青子记账 CSV files")
    parser.add_argument("--output", required=True, help="Output CSV path")
    parser.add_argument("--tickers", default=",".join(DEFAULT_TICKERS), help="Comma-separated ticker list")
    return parser.parse_args()


def latest_cashflow_date(csv_path):
    with Path(csv_path).open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        if not reader.fieldnames or "日期" not in reader.fieldnames:
            raise ValueError("CashFlow CSV is missing 日期")
        dates = []
        for row_number, row in enumerate(reader, start=2):
            value = (row.get("日期") or "").strip()
            if not value:
                continue
            if value == "日期" or (value.startswith("---") and value.endswith("---")):
                continue
            parsed = None
            for date_format in ("%Y-%m-%d", "%m/%d/%Y %H:%M", "%m/%d/%Y"):
                try:
                    parsed = datetime.strptime(value, date_format).date()
                    break
                except ValueError:
                    continue
            if parsed is None:
                raise ValueError(f"Invalid 日期 at CashFlow CSV row {row_number}: {value}")
            dates.append(parsed)
    if not dates:
        raise ValueError("CashFlow CSV has no usable 日期 values")
    latest = max(dates)
    period_match = re.search(r"(\d{4}-\d{2}-\d{2})至(\d{4}-\d{2}-\d{2})", Path(csv_path).name)
    if period_match:
        filename_end = datetime.strptime(period_match.group(2), "%Y-%m-%d").date()
        latest = min(latest, filename_end)
    return latest


def main():
    args = parse_args()
    cashflow_dates = [latest_cashflow_date(csv_path) for csv_path in (args.cashflow_csv or [])]
    cashflow_as_of = max(cashflow_dates) if cashflow_dates else None
    explicit_as_of = datetime.strptime(args.as_of, "%Y-%m-%d").date() if args.as_of else None
    if cashflow_as_of and explicit_as_of and cashflow_as_of != explicit_as_of:
        raise ValueError(f"--as-of {explicit_as_of} does not match latest CashFlow date {cashflow_as_of}")
    as_of = cashflow_as_of or explicit_as_of
    if not as_of:
        raise ValueError("Supply --cashflow-csv or --as-of")
    tickers = list(dict.fromkeys(item.strip().upper() for item in args.tickers.split(",") if item.strip()))
    if not tickers:
        raise ValueError("No tickers supplied")

    start = datetime.strptime(args.start, "%Y-%m-%d").date() if args.start else as_of.replace(day=1)
    if start > as_of:
        raise ValueError(f"--start {start} is after --as-of {as_of}")
    end = as_of + timedelta(days=1)  # yfinance end is exclusive
    download_tickers = [YAHOO_SYMBOLS.get(ticker, ticker) for ticker in tickers]
    if len(set(download_tickers)) != len(download_tickers):
        raise ValueError("Ticker aliases resolve to duplicate Yahoo symbols")

    data = yf.download(
        download_tickers,
        start=start.isoformat(),
        end=end.isoformat(),
        auto_adjust=False,
        progress=False,
        threads=True,
    )
    if data.empty or "Close" not in data:
        raise RuntimeError(f"No closing-price data returned for {start} through {as_of}")

    close = data["Close"]
    if len(download_tickers) == 1:
        close = close.to_frame(name=download_tickers[0])
    close.columns = [str(column).upper() for column in close.columns]
    missing_columns = [ticker for ticker in download_tickers if ticker not in close.columns]
    if missing_columns:
        raise RuntimeError(f"Missing ticker columns: {', '.join(missing_columns)}")
    complete = close[download_tickers].dropna(how="any")
    if complete.empty:
        raise RuntimeError("No common trading date has a valid close for every ticker")

    latest_index = complete.index[-1]
    pricing_date = latest_index.date() if hasattr(latest_index, "date") else date.fromisoformat(str(latest_index)[:10])
    output = Path(args.output).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Ticker", "Current Price", "Currency", "Date"])
        for index, row in complete.iterrows():
            row_date = index.date() if hasattr(index, "date") else date.fromisoformat(str(index)[:10])
            for ticker in tickers:
                yahoo_symbol = YAHOO_SYMBOLS.get(ticker, ticker)
                price = float(row[yahoo_symbol])
                if price <= 0:
                    raise RuntimeError(f"Invalid close for {ticker} on {row_date}: {price}")
                writer.writerow([ticker, f"{price:.6f}", "USD", row_date.isoformat()])

    print(json.dumps({
        "requested_as_of": as_of.isoformat(),
        "as_of_source": "cashflow" if cashflow_as_of else "argument",
        "cashflow_files": args.cashflow_csv or [],
        "pricing_date": pricing_date.isoformat(),
        "first_pricing_date": str(complete.index[0])[:10],
        "trading_days": len(complete.index),
        "tickers": tickers,
        "output": str(output),
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()

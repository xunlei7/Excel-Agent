# Stock mapping

The executable Stock layout, formulas, row limits, and formatting live in `create_workbook.mjs` and `import_stock.mjs`. This reference contains broker interpretation and security metadata.

## Transaction rules
- Broker comes only from the source folder: Fidelity, Robinhood, Charles, or IBKR.
- Normalize Action to exactly `Buy`, `Sell`, or `Dividend`. Robinhood `CDIV` is `Dividend`.
- Dividend uses Units 0 and Trade Price 0 when blank and stores positive dividend cash in `O`; it never changes units.
- Ignore and report ACH, DCF, deposits, withdrawals, and broker cash transfers.
- Preserve explicit Fee and Cost Basis. If Cost Basis is missing, infer it only when the Sell closes the complete position for that Broker + Ticker and the preceding cost roll-forward is complete. Leave partial or otherwise ambiguous sales blank and report them.
- Transaction realized P/L is `Sell Trade Value - Fee - Cost Basis`. Holdings realized P/L sums only Sell realized P/L; Dividend cash is excluded from this column.
- Sort the complete ledger by Date ascending after occurrence-count reconciliation.

Core duplicate key:

`Broker + Date + canonical Ticker + Action + Units + Trade Price + Currency`.

Before reconciliation, combine split executions with identical Broker, Date, Ticker, Action, Trade Price, and Currency. Sum Units, Fee, Cost Basis, and dividend cash.

Fidelity investment-report PDFs use the preceding weekday as inferred trade date for Buy/Sell settlement dates; Dividend keeps its statement date. Ignore Fidelity core money-market income.

## Statement cache

- Keep every historical statement in its broker folder so a template-free workbook can be rebuilt from the full source history.
- `import_stock.mjs` stores content-addressed parse results in `.cache/update-treasury/stock-statements/` by default. A filename change alone does not invalidate a cache entry; a content change, broker-folder change, parser-version change, or invalid cache entry does.
- Cache entries contain normalized parsed transactions and ignored-event reports, never the source statement itself. They are derived data and are safe to delete; the next run recreates them from the statements.
- Use `--refresh-statement-cache` to rebuild all recognized entries, `--no-statement-cache` to parse every file without caching, or `--statement-cache-dir DIR` to choose another cache location.

## Broker columns

Generic/Fidelity aliases include Date or Trade Date, Ticker or Symbol, Name or Security, Action or Side, Units or Quantity, Trade Price or Price, and Amount or Net Amount for dividends.

Charles aliases additionally include `Fees & Comm` and `Fees & Commissions`. IBKR aliases include `Date/Time`, `Buy/Sell`, `T. Price`, `Comm/Fee`, `Basis`, and `Proceeds`.

Robinhood uses `Activity Date`, `Instrument`, `Description`, `Trans Code`, `Quantity`, `Price`, and `Amount`. Discard CUSIP lines from Description. Folder implies Broker Robinhood and Currency USD.

## Canonical securities

`ECHO` maps to `SATS`. Compare tickers case-insensitively. Never use a broker narrative as Name.

| Ticker | Canonical Name |
|---|---|
| AAPL | Apple Inc. |
| AMZN | Amazon.com, Inc. |
| ARM | Arm Holdings plc |
| GOOG | Alphabet Inc. Class C |
| IONQ | IonQ, Inc. |
| MSFT | Microsoft Corporation |
| MSTR | Strategy Inc. |
| NVDA | NVIDIA Corporation |
| QQQ | Invesco QQQ Trust |
| RGTI | Rigetti Computing, Inc. |
| SATS | EchoStar Corporation |
| TSLA | Tesla, Inc. |
| VOO | Vanguard S&P 500 ETF |

## Prices and holdings

- Holdings are calculated from the complete transaction ledger, never copied from a broker snapshot.
- When raw statements do not cover the entire history, begin with the explicit opening Units, Total Cost, and cumulative Realized P/L in `monthly-close-input.md`; add Buy/Sell activity on or after Opening Date. Never treat absent pre-period statements as zero holdings.
- Fetch all modeled prices on the latest complete common trading date on or before the CashFlow cutoff.

## New security

Before importing an unmodeled ticker:

1. Determine canonical Ticker, Name, Currency, Category, and Bucket.
2. Add current-price and holdings rows with adjacent formulas and format.
3. Add a formula-linked Asset row.
4. Add the ticker to both Summary ticker-dependent blocks and extend totals, table, and chart.
5. Import its transactions and prices.

Default ETF to `Growth` and Stock to `High Growth` only when security type is reliable.

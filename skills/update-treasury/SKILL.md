---
name: update-treasury
description: Run and publish the One Piece monthly treasury close while preserving workbook formulas, formatting, charts, and structure. Use when Codex needs to start a workbook from code, continue from the verified bootstrap, import cash-flow exports or Fidelity, Robinhood, Charles, and IBKR activity, reconcile transactions and balances, update Asset, Stock, prices, and Summary, or publish one monthly archive. Trigger for start, continue, One Piece, treasury-close, cash-flow-import, brokerage-import, balance-update, or price-update requests.
---

# Update Treasury

Use this Skill for `start`, `continue`, `One Piece`, `小金库`, or requests to import monthly cash-flow, brokerage, balance, holding, or price data into the Wealth workbook.

## Required context

1. Read the spreadsheet Skill and load its bundled workspace dependencies. In a Terminal session started by `./start-codex.sh`, `CODEX_NODE_MODULES` is the preloaded, loader-provided spreadsheet runtime and `WEALTH_PYTHON` is the validated Python runtime used only for price fetching; continue with them even when the CLI does not expose a `load_workspace_dependencies` tool.
2. Read [references/mapping.md](references/mapping.md), then read only the sheet references required by the request. Read every linked reference for a complete monthly close.
3. Read `monthly-close-input.md` from the project root. Values supplied directly by the user take precedence over that file. Treat its opening cash and opening-holdings table as the explicit pre-period baseline; do not assume missing historical CSVs or statements imply zero opening positions.
4. For a complete close, run `scripts/update_treasury.mjs`. It calls the starting-workbook step, `import_cashflow.mjs`, `import_stock.mjs`, `update_asset.mjs`, `update_summary.mjs`, and `apply_workbook_format.mjs` in that order. `start` builds from code. `continue` explicitly loads `Output/bootstrap/One_Piece.xlsx`. Do not run standalone scripts separately for a complete close.

The public [assets/One_Piece_Template.xlsx](assets/One_Piece_Template.xlsx) workbook is a privacy-safe example of the code-built four-sheet structure and presentation. It contains only zero-value sample inputs. Do not use it as an opening-balance source or as a substitute for `monthly-close-input.md` or the private Bootstrap used by `continue`.

## Monthly-close workflow

### 1. Import CashFlow

- Run `scripts/import_cashflow.mjs` for CashFlow-only updates. The main `scripts/update_treasury.mjs` entry point imports and reuses the same implementation.
- For fast visual testing, pass `--cashflow-only`. It creates a temporary workbook from code, runs only the CashFlow importer, and creates a standalone workbook containing one sheet. This preview must not be used as a monthly-close archive.
- Read every recognized 青子记账 CSV in `CashFlow/` and normalize whitespace, dates, currencies, and account aliases.
- In a complete close, use `As-of Date` from `monthly-close-input.md` as the default CashFlow cutoff. Report later source rows as excluded; do not let them move the workbook into the next close month. An explicit `--through-date` overrides this cutoff.
- Split records by `类型` and process in this order: `Income`, `Expense`, `Transfer`.
- Sort each block by date ascending, then append it in the matching CashFlow block using that block's formulas and formatting.
- Reconcile duplicates by occurrence count so accumulated CSV files do not re-import old records and legitimate identical transactions are retained.
- Transfer rows require a reliable Counterparty. Use the mapping reference; if the required source field is blank, stop and identify the rows instead of guessing.
- Reconcile every source row to exactly one of: imported, skipped as an existing occurrence, or rejected with a stated reason.

### 2. Import brokerage activity and update Stock

- Run `scripts/import_stock.mjs` for Stock-only development and testing. Pass `--stock-only` to create a workbook containing only the Stock sheet; this preview is not a monthly-close archive.
- Read supported monthly statements from `Fidelity/`, `Robinhood/`, `Charles/`, and `IBKR/`. Treat the source folder as authoritative for Broker and default a missing currency to USD; ignore any conflicting Broker field inside a statement. Fidelity PDF statements and recognized brokerage CSV exports are supported; unsupported files must be reported.
- Keep historical statements in their broker folders. Cache parsed statement results by file content in `.cache/update-treasury/stock-statements/`; unchanged files must use the cache, while new, changed, corrupt-cache, or parser-version-mismatched files must be parsed again. The cache is derived data and may be deleted without losing source records.
- Import only security events into `Stock!交易明细`; report deposits, withdrawals, ACH, cash transfers, and unsupported events separately.
- Normalize the Action written to the workbook to exactly `Buy`, `Sell`, or `Dividend`. Robinhood `CDIV` is `Dividend`. Do not write `Dividend|`; Summary formulas use `Dividend`.
- Normalize every Stock transaction and holdings Name from its canonical Ticker. Never place broker activity narratives such as `Cash Div: R/D ...` in the Name column.
- A Buy increases units, a Sell decreases units, and a Dividend changes cash but never units. Preserve explicit fees and cost basis; never invent missing values.
- Preserve a statement-provided Cost Basis. When it is missing, infer it only if the sale closes the complete broker-and-ticker position and the full cost roll-forward is known; otherwise leave it blank and report it. Realized P/L equals Sell Trade Value minus Fee minus Cost Basis. Holdings realized P/L includes Sell rows only and never Dividend cash.
- Reconcile duplicates by occurrence count, then sort the complete Stock transaction ledger—not only newly accepted rows—by trade/activity date ascending from January through December.
- Update `持仓汇总（按当前市值）` from the complete transaction ledger, not from the brokerage CSV's ending snapshot.
- Start holdings, total cost, and cumulative realized P/L from the explicit opening-holdings table, then apply only transactions on or after Opening Date. This baseline replaces earlier activity that is not available as raw statements.
- Update `当前价格输入（可手动更新）` by calling `scripts/fetch_stock_prices.py` for every modeled ticker and using the latest complete common price date. In Stock-only mode, derive the cutoff from the latest recognized CashFlow CSV date unless `--price-as-of` is supplied; use `--skip-price-fetch` only for an explicit offline test.

#### New stock or ETF

When a brokerage CSV contains a ticker not already modeled, use the structural-update path before importing its transactions:

1. Establish canonical ticker, name, currency, security type, and investment bucket. Use explicit source metadata. If type is clear but bucket is omitted, default ETF to `Growth` and Stock to `High Growth`, and report the assumption; otherwise ask the user.
2. Add the ticker to Stock current-price input and holdings summary, copying the adjacent formulas and format.
3. Add it to Asset with a formula-linked market value; never hardcode a security's Amount.
4. Add it to both Summary ticker blocks—Linked Holdings/prices and real-time Units Held—in the same order. Extend dependent totals, history formulas, table ranges, and chart ranges.
5. Complete the structural extension, then import the new security's Buy/Sell/Dividend rows.

Do not silently drop a new ticker or squeeze it into an unrelated cell. If the existing layout lacks capacity and cannot be safely extended on a working copy, stop and report the exact constraint.

### 3. Update Asset

- Run `scripts/update_asset.mjs` after Stock is current. It reads the reusable balance table in `monthly-close-input.md` and keeps the Stock dependency sheet in the output workbook.
- Match reported cash/account balances by `资产名称 + 币种`; update only Amount and Update time for those rows.
- Keep omitted balances unchanged. Never interpret omission as zero.
- Ensure every held stock or ETF has one Asset row whose Amount is linked to Stock holdings market value, with the correct currency, Category, Bucket, FX formula, and converted-amount formula.
- Brokerage cash accounts remain separate Cash assets and are updated only from the user's reported balances.

### 4. Update Summary

- Run `scripts/update_summary.mjs` for a standalone Summary update after CashFlow and Stock are current. Supply `--price-csv` for an offline/reproducible test; otherwise the script calls `scripts/fetch_stock_prices.py`. This script keeps the dependency sheets because Summary formulas reference CashFlow, Stock, and Asset.
- Keep Linked Holdings synchronized with the modeled ticker list and recompute current Units Held from Buy minus Sell transactions.
- Keep Summary row 12 blank in the Linked Holdings/price area. Security names remain in Stock and are not repeated in Summary.
- For each Summary history date, calculate historical Units Held as of that date and live cash from opening cash, CashFlow, Buy, Sell, and Dividend activity.
- Start Summary units and live cash from the explicit opening baseline. Do not rebuild a pre-period balance from incomplete source history.
- Use the latest recognized CashFlow date to choose the month and requested pricing cutoff. Fetch one complete set of unadjusted closes for every modeled ticker on every common trading day from the first day of that month through the cutoff. Do not create weekend or market-holiday rows.
- Add Summary dates in ascending order and never overwrite an existing date with different prices.
- Recalculate Stock MV (USD), Stock MV (Base), Total Net Worth (Base), and the latest-value KPIs from the latest populated Summary date.

### 5. Publish

- Run `scripts/apply_workbook_format.mjs` as the final stage. It deterministically reapplies four independent, sheet-specific style systems derived from `Output/bootstrap/One_Piece.xlsx`, including each sheet's titles, headers, fonts, alignments, data-row borders, widths, colors, and number formats. In `start` the Bootstrap is a visual reference only; in explicit `continue` it is the starting workbook.
- Write the single completed result to `Output/YYYY-MM/One_Piece_YYYY-MM.xlsx`, where `YYYY-MM` comes from the latest recognized CashFlow date. Do not create a duplicate latest-copy in `Output/` or at the project root.
- Report CashFlow counts by Income/Expense/Transfer, brokerage counts by Buy/Sell/Dividend, ignored broker events, new securities, balance changes, price dates, Summary rows, and output paths.

## Preferred use

To build a new workbook from code, use:

```text
start
```

To continue from the verified Bootstrap workbook, use:

```text
continue
```

The Skill reads the folders and `monthly-close-input.md`. The user may instead report balances directly in the prompt. Convert only explicitly listed, nonblank balances to the temporary asset snapshot.

The complete-close command is:

```bash
node skills/update-treasury/scripts/update_treasury.mjs --mode start
```

For Bootstrap continuation, use `--mode continue`. Its default source is `Output/bootstrap/One_Piece.xlsx`; override it only with `--bootstrap-workbook FILE`.

The orchestrator creates intermediate workbooks in a temporary directory and deletes them after the run. Only the completed monthly archive is published. Refuse to replace an existing archive unless replacement is explicitly authorized with `--allow-existing-output`.

## Safety boundaries

- Keep dates and numbers typed, not text. Preserve the workbook's established formats, formulas, charts, tables, and the sheet order `CashFlow`, `Stock`, `Asset`, `Summary`. Do not create a helper or dropdown-options sheet.
- Never write beyond a verified data block. Extend a structural block only when a new security requires it and every dependency is updated together.
- Do not classify a CSV by filename alone. Check its headers and broker-folder context.
- Do not infer missing transfer counterparties, trade actions, tickers, quantities, prices, dividend cash amounts, currencies, or conflicting broker identities.
- Treat price history as an external source: require a positive close for every modeled ticker on each accepted trading day.
- Preserve ticker aliases defined in the mapping reference consistently across transactions, prices, holdings, Asset, and Summary.

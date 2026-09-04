# Workbook and input conventions

## Paths and publication

- `start`: build a new workbook from code and raw inputs.
- `continue`: explicitly begin with `Output/bootstrap/One_Piece.xlsx`, then run the same modules and remove every non-required sheet.
- Publish exactly one monthly result at `Output/YYYY-MM/One_Piece_YYYY-MM.xlsx`. Do not create another latest-copy.
- Derive `YYYY-MM` from the latest accepted 青子记账 date.

Preserve exactly four sheets in this order: `CashFlow`, `Stock`, `Asset`, `Summary`. Do not create a helper or dropdown-options sheet.

## Template-free bootstrap inputs

Read these values from `monthly-close-input.md`; never recover them from a prior workbook:

- Opening Date.
- Opening Cash Balance (Base).
- Opening holdings: Ticker, Units, Total Cost, cumulative Realized P/L, and Bucket.
- Base Currency.
- USD/CNY exchange rate.
- As-of Date, which is the default cutoff for a complete close so later transactions already present in an accumulated export are not imported early.

Opening Date is the first date whose activity is applied to the explicit opening cash and holdings. The current baseline is the verified 2026-07-31 bootstrap close, so Opening Date is 2026-08-01. This preserves the correct history in `start`; the Bootstrap workbook is loaded only for explicit `continue`.

`scripts/create_workbook.mjs` owns the code-built four-sheet layout and the explicit Bootstrap-continuation loader. `scripts/apply_workbook_format.mjs` owns the deterministic presentation layer and is reapplied after all data modules. Its rules are derived from `Output/bootstrap/One_Piece.xlsx`; that workbook is loaded only for `continue`, never for `start`. See [formatting.md](formatting.md) for the presentation invariants.

`scripts/update_treasury.mjs` selects exactly one requested mode, then passes the starting workbook through CashFlow, Stock, Asset, Summary, and final formatting in that order. It must never silently switch modes. Intermediate workbooks are temporary; only the monthly archive is published.

## Input discovery

Classify files by normalized headers plus folder context, not filename alone.

- `CashFlow/`: 青子记账 exports with `日期, 分类, 备注, 类型, 金额, 货币, 账户`.
- `Fidelity/`, `Robinhood/`, `Charles/`, and `IBKR/`: supported statements. The folder is authoritative for Broker and defaults a missing Currency to USD.
- Project root: optional price CSV or asset snapshot CSV.

Allow accumulated exports. Reconcile duplicate occurrences and report every ignored, unsupported, or rejected file and row.

## Shared workbook rules

- For `start`, create the workbook structure from code. For `continue`, use only the explicitly selected Bootstrap workbook as the starting point.
- Preserve formulas, cell types, formats, tables, charts, hidden-sheet state, and sheet order.
- Use typed dates and numbers rather than formatted strings.
- Do not write beyond verified data blocks.
- Normalize exported shared formulas to ordinary cell formulas when the writer retains stale shared-formula records; the final workbook must open in Excel without a repair prompt.

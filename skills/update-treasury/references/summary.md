# Summary mapping, formulas, and format

The executable Summary layout, formulas, table, chart, and formatting live in `create_workbook.mjs` and `update_summary.mjs`. This reference contains the financial meaning of the calculations.

## Live cash

For date d:

`Opening Cash + Income - Expense - Buy Trade Value + Sell Proceeds + Dividend Cash`.

- Apply every term only from the opening date through d.
- Opening Cash is the verified balance immediately before Opening Date. Opening security Units come from the opening-holdings table; transactions on or after Opening Date roll both balances forward.
- Detect CashFlow base-amount columns by row-11 headers. Compact CashFlow uses `J/W` (`Amount (Base)`); the legacy template uses `L/Y` (`Converted Amount`). Never hardcode one layout for Summary-only updates.
- Transfers and brokerage ACH/DCF movements do not change consolidated live cash.
- Use the Stock transaction Date, Action, Trade Value, and Dividend cash columns.
- Total Net Worth (Base) equals Stock MV (Base) plus live cash.

## Daily history

- Pricing history runs from Opening Date through the latest accepted CashFlow date.
- Require a positive unadjusted USD close for every modeled ticker on each common trading day.
- Keep fully sold tickers in the historical price and units matrix even though they no longer appear in current holdings.
- Do not add weekends or market holidays.
- Preserve existing dates and prices; append or merge dates in ascending order and reject conflicting prices.
- In `continue`, preserve verified Bootstrap history before Opening Date as fixed historical snapshots. Freeze it before the first intermediate export so unsupported legacy formulas cannot replace valid historical cash and net-worth cached values. Apply the opening baseline and current formulas only from Opening Date forward.
- Historical Units Held on d equals Buy units through d minus Sell units through d.
- Update the table bounds, net-worth chart ranges, and latest-value KPI formulas to the dynamic width and last populated date.

## Reconciliation and opening anchor

- Keep `N3:O3` as the merged `Reconciliation Check` title.
- Keep labels in `N4:N7`: Real Asset, Calculated Asset, Total Difference, and Status.
- Keep formulas and results in `O4:O7`: Asset total, latest calculated net worth, difference, and tolerance status.
- Style the dynamic Opening Date and Opening Cash Balance label cells with the same dark-blue fill and white bold Calibri 11 font as the Reconciliation Check title.

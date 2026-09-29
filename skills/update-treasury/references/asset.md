# Asset mapping

The executable Asset layout, formulas, and formatting live in `create_workbook.mjs` and `update_asset.mjs`. This reference contains balance and classification rules.

- Run `scripts/update_asset.mjs` after Stock is current. It reads the balance table in `monthly-close-input.md`; a nonblank `As-of Date` or `--asset-date YYYY-MM-DD` is required.
- Match cash/account rows by `资产名称 + 币种`.
- Update only Amount and Update time for explicitly reported balances. Keep omitted balances unchanged and never interpret omission as zero.
- Brokerage cash remains separate Category `Cash` assets.
- Keep the account rows in this order before securities: WeChat, 中信银行 USD, 中信银行 CNY, 支付宝, Chase, Cash, Robinhood Cash, Fidelity Cash, IBKR Cash, Charles Cash. In particular, IBKR Cash and Charles Cash sit immediately after Fidelity Cash.
- Synchronize every modeled ticker in `Stock!P22:Z44` to exactly one Asset Stock/ETF row. Add a missing ticker to the first available Asset detail row and copy the existing security-row format.
- A stock or ETF row uses canonical ticker/name, Category, Currency, formula-linked Amount from the exact Stock holdings Current Value cell, FX formula, Converted Amount formula, Bucket, and formula-linked Stock price update date.
- Existing label `TESLA` corresponds to `TSLA`; preserve it unless the template is deliberately normalized.
- Never hardcode a stock or ETF Amount.
- Convert every asset to the configured Base Currency using the explicit USD/CNY input.
- Keep the Bootstrap-style FX panel complete: base currency, USD/CNY, CNY and USD conversion factors, status, and explanatory labels. Apply visible cell borders to every Asset detail row.
- Center every cell in the `Asset details` header and detail grid horizontally and vertically.
- Display every populated `Update time` value in blue, including Stock-linked dates.
- Keep a Stock-linked `Update time` blank when its source price date is blank; never display an Excel zero date such as `1900-01-00`.

Summary live cash is a transaction roll-forward from its dated opening cash anchor. Asset cash balances provide a current reconciliation target but must not be substituted into historical Summary rows without a dated anchor and a complete reverse roll-forward.

# CashFlow mapping and format

The executable block layout, formulas, and formatting live in `create_workbook.mjs` and `import_cashflow.mjs`. This reference contains only source interpretation and business mapping.

Process Income, Expense, and Transfer in that order. Sort each accepted block by Date ascending. Reconcile accumulated CSVs by occurrence count.

Book and Note/Notes are not part of the compact blocks. Never import `账本`, `报销`, `Book`, `Note`, or `Notes`.

## 青子记账 mapping

| 青子记账 | Income | Expense | Transfer |
|---|---|---|---|
| 日期 | Date | Date | Date |
| 日期 YYYY-MM | Year-Month | Year-Month | Year-Month |
| 备注 | Category, fallback 分类 | Tag | Counterparty; required |
| 分类 | fallback Category | Category | Category |
| 详细备注 | Description | Description | Description |
| 类型 | 收入 | 支出 | 转账, 转入, 转出 |
| 金额 | Amount | Amount | signed Amount |
| 货币 | Currency | Currency | Currency |
| 账户 | Account | Account | Account |

Transfer signs:

- `转入`: positive absolute Amount.
- `转出`: negative absolute Amount.
- `转账`: preserve the source sign.
- Never infer Counterparty from Description or filename.

For a separate `---转账记录---` section, write balanced negative and positive rows only for a supported same-currency account pair. Use the explicit account pair as Counterparty when the note is blank. Stop for unknown accounts or cross-currency transfers.

## Account aliases

- `微信钱包`, `微信`, `Wechat` -> `WeChat`.
- Preserve `Chase`, `中信银行`, `支付宝`, `Cash`, `Robinhood Cash`, `Fidelity Cash`, `IBKR Cash`, and `Charles Cash`.
- Supported currencies: USD and CNY.

Reconciliation keys exclude Book, Note/Notes, and formula outputs. Verify `source rows = imported + skipped existing occurrences + rejected rows` for each block.

## Fixed summary and table presentation

- Rewrite the fixed labels on every import and safe export: `Selected Month`, the three total labels, Start Date, End Date, record counts, Net Cash Flow, and every cell in all three table header rows. The three `Year-Month` headers must remain at `B11`, `O11`, and `AB11`.
- `B2` uses Cambria 11 with a visible light-blue fill.
- `E3:E5` and `H3:H5` are first-date and last-date formulas for Income, Expense, and Transfer. Row 6 stays blank because there are only three CashFlow blocks.
- Apply thin borders to every populated table cell, every table header, and each fixed summary block.

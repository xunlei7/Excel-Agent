# CashFlow sources, mapping, and format

The executable block layout, formulas, and formatting live in `create_workbook.mjs`, `import_cashflow.mjs`, and `cashflow_sources.mjs`. This reference contains source interpretation and business rules.

## Source folders

The immediate folder under `CashFlow/` determines the account. Do not read an account identity from a filename or statement field.

| Folder | Workbook account | Currency | Supported source |
|---|---|---|---|
| `Chase_Checking_<last4>` | Chase Checking + last four digits | USD | Chase activity CSV |
| `Chase_Credit_<last4>` | Chase Credit + last four digits | USD | Chase activity CSV |
| `Alipay` | 支付宝 | CNY | Alipay CSV |
| `WeChat` | WeChat | CNY | WeChat XLSX |
| `CITIC_CNY` | 中信银行 | CNY | CITIC XLS/XLSX |
| `CITIC_USD` | 中信银行 | USD | CITIC XLS/XLSX |

Files may accumulate in these folders. Apply the As-of Date before classification and report excluded later rows. Preserve workbook rows through 2026-07-31 and replace rows on or after 2026-08-01 from the account sources. Root-level 青子记账 files are retained only as private history and are not part of the new-period import.

The importer calculates a SHA-256 content hash for every source. Parsed transactions are cached in `.cache/update-treasury/cashflow-sources/`; renaming an unchanged file can reuse the same cache entry. A content, account-folder, parser, or currency change creates a different cache key. Use `--no-cashflow-cache` for a one-time uncached run or `--refresh-cashflow-cache` to rebuild all discovered entries.

For overlapping exports, build a normalized key from Account, Currency, Date, signed Amount, Type, Description, source Category, and direction. Preserve the largest occurrence count present in any single file. This removes overlap between accumulated files without collapsing two genuinely identical transactions contained in one statement.

## Field mapping

- `Date`: source transaction date, not export date.
- `Year-Month`: derived from Date.
- `Category`: Chase Credit category after the fixed category map. Checking, WeChat, Alipay, and CITIC remain blank unless an explicit reusable rule supplies a category.
- `Type`: the source Type or transaction-type field without semantic replacement.
- `Description`: merchant or source description. A private rule may replace it with a stable description that contains no unnecessary personal detail. Remove date tokens already represented by the Date column.
- `Amount`: positive Income, positive Expense, negative Expense for refunds/reimbursements, and signed Transfer from the perspective of the row's Account.
- `Currency` and `Account`: determined by the source folder.

Chase Credit `Sale` is Expense, `Return` is negative Expense, and `Payment` is an ignored settlement. Chase Checking debit-card purchases are Expense. Credit-card payments are ignored. Brokerage movements and ATM cash deposits are balanced internal Transfer rows. Credit-card accounts never become Asset rows; the Asset `Chase` balance remains the checking balance supplied in `monthly-close-input.md`.

## Ambiguous transactions

Read `CashFlow/cashflow-rules.csv` before applying automatic rules. Match by source folder plus the supplied date, signed amount, type, and optional description fragment. Stop with an exact unresolved-row report when a personal transfer cannot be classified. After confirmation, save the decision in the private rule file so later runs remain deterministic.

Cross-currency transfer legs share one Counterparty key. Store the confirmed transaction-specific conversion rate on the non-base-currency leg so Outstanding rounds to zero. Same-currency reversals and internal transfers also share one Counterparty key and must settle to zero.

Do not append a transaction date to Counterparty or Description. The private rule's `date` column remains part of matching and reconciliation, but it is not copied into either display field.

## Fixed layout and presentation

- Process and sort Income, Expense, and Transfer by Date ascending.
- Headers are `Type`, never `Tag`, at `D11`, `Q11`, and `AE11`.
- Book and Note/Notes are not part of the compact blocks.
- Rewrite the fixed labels and all three header rows on every import and safe export.
- `B2` uses Cambria 11 with a visible light-blue fill.
- `E3:E5` and `H3:H5` show first and last dates; row 6 stays blank.
- Apply thin borders to every populated table cell, table header, and fixed summary block.

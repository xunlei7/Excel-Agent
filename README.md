# Excel Agent

Excel Agent is a Codex skill and script collection for building a four-sheet monthly treasury workbook (`CashFlow`, `Stock`, `Asset`, and `Summary`) from cash-flow exports, brokerage statements, and manually supplied balances.

A privacy-safe blank workbook is available at `skills/update-treasury/assets/One_Piece_Template.xlsx`. It demonstrates the generated layout and formatting but is not an opening-balance source.

## Private data

The repository intentionally excludes all local financial data and generated workbooks:

- `CashFlow/`, including account-export subfolders and the private classification-rule file
- `Fidelity/`, `Robinhood/`, `Charles/`, and `IBKR/`
- `Output/`
- `monthly-close-input.md`
- caches, virtual environments, and workbook-analysis artifacts

These paths are protected by `.gitignore`. Do not force-add their contents.

## Local setup

1. Copy `monthly-close-input.example.md` to `monthly-close-input.md` and enter private bootstrap values, the close date, and account balances.
2. Under `CashFlow/`, create `Chase_Checking_<last4>`, one `Chase_Credit_<last4>` folder per card, `Alipay`, `WeChat`, `CITIC_CNY`, and `CITIC_USD`, then place each export in its source folder.
3. Install the Python dependencies from `skills/update-treasury/requirements.txt` in the environment selected by `WEALTH_PYTHON`.
4. Start Codex with `./start-codex.sh`, then request `start` for a code-built workbook or `continue` to use a local bootstrap workbook.

The completed workbook is written only to `Output/YYYY-MM/One_Piece_YYYY-MM.xlsx` and remains local.

CashFlow and brokerage parsers keep content-addressed caches under `.cache/update-treasury/`. Original exports remain the source of truth; the cache can be deleted and rebuilt at any time.

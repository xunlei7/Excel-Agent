# Excel Agent

Excel Agent is a Codex skill and script collection for building a four-sheet monthly treasury workbook (`CashFlow`, `Stock`, `Asset`, and `Summary`) from cash-flow exports, brokerage statements, and manually supplied balances.

## Private data

The repository intentionally excludes all local financial data and generated workbooks:

- `CashFlow/`
- `Fidelity/`, `Robinhood/`, `Charles/`, and `IBKR/`
- `Output/`
- `monthly-close-input.md`
- caches, virtual environments, and workbook-analysis artifacts

These paths are protected by `.gitignore`. Do not force-add their contents.

## Local setup

1. Copy `monthly-close-input.example.md` to `monthly-close-input.md` and enter private bootstrap values, the close date, and account balances.
2. Create the input directories listed above and place each export or statement in the folder for its source.
3. Install the Python price dependency from `requirements.txt` in a suitable environment.
4. Start Codex with `./start-codex.sh`, then request `start` for a code-built workbook or `continue` to use a local bootstrap workbook.

The completed workbook is written only to `Output/YYYY-MM/One_Piece_YYYY-MM.xlsx` and remains local.

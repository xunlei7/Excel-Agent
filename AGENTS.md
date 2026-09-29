# Wealth

- When running from Terminal, start Codex with `./start-codex.sh`. It exports the loader-provided spreadsheet runtime through `CODEX_NODE_MODULES` and the Conda `money` price-fetching runtime through `WEALTH_PYTHON`. If those variables are present and their paths validate, use them; the absence of a `load_workspace_dependencies` tool is not by itself a blocker.
- For `start`, `continue`, `One Piece`, or `小金库` requests, follow `skills/update-treasury/SKILL.md`, read `monthly-close-input.md`, import account exports from the supported source folders under `CashFlow/`, and import supported brokerage statements from `Fidelity/`, `Robinhood/`, `Charles/`, and `IBKR/`.
- Fetch Summary prices using the latest CashFlow date when no price CSV is supplied.
- Write each completed close to `Output/YYYY-MM/One_Piece_YYYY-MM.xlsx`, where `YYYY-MM` comes from the latest recognized CashFlow date.
- Treat that monthly archive as the only published result. Do not create a duplicate latest-copy in `Output/` or at the project root.
- `start` must build the four-sheet workbook from code and raw inputs. `continue` must explicitly begin with `Output/bootstrap/One_Piece.xlsx`; never switch modes silently. Both modes publish only the four required sheets.

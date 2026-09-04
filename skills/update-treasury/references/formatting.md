# Deterministic workbook formatting

`scripts/apply_workbook_format.mjs` is the single executable owner of the fixed presentation layer for all four sheets. Each sheet has an independent formatting function and style system; shared helpers may only apply low-level operations such as a border or merged band. Run the formatter after every data-writing module so imports cannot leave titles, headers, fonts, fills, borders, widths, row heights, alignments, or number formats in a partially changed state.

The style baseline is `Output/bootstrap/One_Piece.xlsx`. `start` uses it only for development and visual comparison; explicit `continue` uses it as the starting workbook before the deterministic formatter is reapplied.

Required invariants:

- CashFlow uses its own compact cash-flow style: Noto Sans CJK SC labels, Cambria summary values, green imported data, dark-blue block bands, a merged `A1:L1` Summary banner, and thin borders around every populated table cell. The Selected Month value uses blue Cambria 11 text on a light-blue fill.
- Stock uses its own analytical-table style: Noto Sans CJK SC section titles, Cambria transaction and holding data, a light-blue transaction band, a green holdings band, and visible borders for transaction, price, and holdings rows. Clear orphaned Bootstrap header fills outside the current blocks, including `AA21`.
- Asset uses its own source-role style: blue manual balances, green Stock-linked values, gray formula columns, pale-blue centered detail rows, a purple FX input band, and visible borders on every reserved detail row.
- Summary uses its own trend style: Calibri title and table text, dark-blue summary and history headers, light-blue KPI labels, `N4:N7` reconciliation labels with `O4:O7` results, dark-blue opening-anchor labels with white text, and dynamically expanding ticker-dependent price and units sections.
- Formatting code may restore fixed labels and presentation structure, but it must not replace business formulas, imported values, charts, or dynamic row contents.

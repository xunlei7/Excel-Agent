#!/usr/bin/env node

import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { execFile } from "node:child_process";
import { createRequire } from "node:module";
import { fileURLToPath, pathToFileURL } from "node:url";
import { promisify } from "node:util";
import { discoverCashflowCsvs } from "./import_cashflow.mjs";
import { canonicalSecurityName } from "./import_stock.mjs";
import { readBootstrap } from "./create_workbook.mjs";
import { restoreFixedTextForSheet } from "./workbook_io.mjs";

const dependencies = process.env.CODEX_NODE_MODULES;
if (!dependencies) throw new Error("CODEX_NODE_MODULES is required");
const require = createRequire(path.join(dependencies, "package.json"));
const { FileBlob, SpreadsheetFile, Workbook } = require("@oai/artifact-tool");
const JSZip = require("jszip");
const execFileAsync = promisify(execFile);

const FIRST_DATA_ROW = 14;
const MAX_DATA_ROW = 1046;
const STOCK_FIRST_ROW = 21;
const STOCK_LAST_ROW = 1004;
const HOLDINGS_FIRST_ROW = 38;
const HOLDINGS_LAST_ROW = 67;
const SUMMARY_HEADERS = ["Stock MV (USD)", "Stock MV (Base)", "Total Net Worth (Base)"];
const DEFAULT_PRICE_SCRIPT = fileURLToPath(new URL("./fetch_stock_prices.py", import.meta.url));

export async function runSummaryUpdate({
  workbookPath,
  outputPath,
  inputDir,
  inputFile = null,
  priceCsv = null,
  summaryDate = null,
  dryRun = false,
  allowExistingOutput = false,
}) {
  const sourceWorkbook = path.resolve(workbookPath);
  const destinationWorkbook = path.resolve(outputPath);
  if (sourceWorkbook === destinationWorkbook) fail("Refusing to overwrite the source workbook. Choose a distinct output path.");
  if (!dryRun && !allowExistingOutput && await exists(destinationWorkbook)) fail(`Output already exists: ${destinationWorkbook}`);

  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(sourceWorkbook));
  const bootstrap = await readBootstrap(path.resolve(inputFile || path.join(inputDir, "monthly-close-input.md")));
  const holdings = readStockHoldings(workbook.worksheets.getItem("Stock"));
  if (!holdings.length) fail(`Stock holdings summary has no tickers at Stock!P${HOLDINGS_FIRST_ROW}:P${HOLDINGS_LAST_ROW}.`);
  const summarySheet = workbook.worksheets.getItem("Summary");
  const openingDate = excelDateText(readOpeningAnchor(summarySheet, detectLayout(summarySheet)).date);

  const priceSource = priceCsv
    ? { file: path.resolve(priceCsv), fetched: false, fetchReport: null }
    : await fetchMonthlyPrices(inputDir, holdings.map((item) => item.ticker), summaryDate, openingDate);
  const history = await readPriceHistory(priceSource.file);
  const requestedDate = resolveRequestedDate(summaryDate, priceSource.fetchReport, history);
  const report = updateSummarySheet(workbook, history, holdings, requestedDate, bootstrap);

  if (!dryRun) {
    await fs.mkdir(path.dirname(destinationWorkbook), { recursive: true });
    await saveSummaryWorkbook(workbook, destinationWorkbook, report);
  }
  return {
    workbook: sourceWorkbook,
    output: dryRun ? null : destinationWorkbook,
    priceSource: priceSource.file,
    fetchedPrices: priceSource.fetched,
    fetchReport: priceSource.fetchReport,
    requestedDate,
    ...report,
  };
}

async function saveSummaryWorkbook(workbook, workbookPath, report) {
  const expectedFormulas = collectExpectedFormulas(workbook, report);
  const exported = await SpreadsheetFile.exportXlsx(workbook);
  const archive = await JSZip.loadAsync(Buffer.from(exported.data));
  const workbookXml = await archive.file("xl/workbook.xml")?.async("string");
  const relationshipsXml = await archive.file("xl/_rels/workbook.xml.rels")?.async("string");
  if (!workbookXml || !relationshipsXml) fail("Exported workbook is missing its workbook metadata.");

  const relationshipTargets = new Map(
    [...relationshipsXml.matchAll(/<Relationship\b[^>]*>/g)].map((match) => match[0]).map((tag) => [
      tag.match(/\bId="([^"]+)"/)?.[1],
      tag.match(/\bTarget="([^"]+)"/)?.[1],
    ]),
  );
  for (const sheetTag of workbookXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sheet\b[^>]*>/g)) {
    const name = sheetTag[0].match(/\bname="([^"]+)"/)?.[1];
    const relationshipId = sheetTag[0].match(/\br:id="([^"]+)"/)?.[1];
    const target = relationshipTargets.get(relationshipId);
    if (!name || !target) continue;
    const worksheetPath = target.startsWith("/") ? target.slice(1) : path.posix.normalize(path.posix.join("xl", target));
    const worksheetFile = archive.file(worksheetPath);
    if (!worksheetFile) fail(`Exported workbook is missing ${worksheetPath}.`);
    const formulaMap = expectedFormulas.get(name) || new Map();
    let worksheetXml = await worksheetFile.async("string");
    worksheetXml = worksheetXml.replace(
      /<(?:[A-Za-z_][\w.-]*:)?c\b(?=[^>]*\br="([A-Z]+\d+)")[^>]*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?c>)/g,
      (cellXml, address) => normalizeCellFormulaXml(cellXml, formulaMap.get(address)),
    );
    worksheetXml = restoreFixedTextForSheet(worksheetXml, name);
    archive.file(worksheetPath, worksheetXml);
  }
  await fs.writeFile(workbookPath, await archive.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

function collectExpectedFormulas(workbook, report) {
  const result = new Map();
  for (let index = 0; ; index += 1) {
    let sheet;
    try { sheet = workbook.worksheets.getItemAt(index); } catch { break; }
    if (!sheet) break;
    const used = sheet.getUsedRange();
    const formulas = used?.formulas || [];
    const start = parseCellAddress(String(used?.address || "A1").split(":")[0]);
    const formulaMap = new Map();
    formulas.forEach((row, rowOffset) => row.forEach((formula, colOffset) => {
      if (typeof formula !== "string" || !formula.startsWith("=")) return;
      const rowNumber = start.row + rowOffset;
      const colIndex = start.col + colOffset;
      if (sheet.name === "Summary" && !isExpectedSummaryFormula(rowNumber, colIndex, report)) return;
      formulaMap.set(`${columnLetter(colIndex)}${rowNumber}`, formula);
    }));
    result.set(sheet.name, formulaMap);
  }
  return result;
}

function isExpectedSummaryFormula(row, col, report) {
  if (row < 10) return true;
  const priceEnd = columnIndex(report.layout.priceEndCol);
  const metricStart = columnIndex(report.layout.metricStartCol);
  const totalEnd = columnIndex(report.layout.totalNetWorthCol);
  const helperStart = columnIndex(report.layout.helperStartCol);
  const helperEnd = columnIndex(report.layout.helperEndCol);
  const cashCol = columnIndex(report.layout.cashCol);
  const lastHistoryRow = FIRST_DATA_ROW + report.totalHistoryRows - 1;
  if (row === 11) return col >= 1 && col <= priceEnd;
  if (row < FIRST_DATA_ROW || row > lastHistoryRow) return false;
  return (col >= metricStart && col <= totalEnd) || (col >= helperStart && col <= helperEnd) || col === cashCol;
}

function normalizeCellFormulaXml(cellXml, expectedFormula) {
  const formulaPattern = /<(?:[A-Za-z_][\w.-]*:)?f\b[^>]*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?f>)/g;
  const withoutFormula = cellXml.replace(formulaPattern, "");
  if (!expectedFormula) return withoutFormula;
  const prefix = cellXml.match(/^<([A-Za-z_][\w.-]*:)?c\b/)?.[1] || "";
  const formulaXml = `<${prefix}f>${escapeXml(expectedFormula.slice(1))}</${prefix}f>`;
  if (/\/>$/.test(withoutFormula)) return withoutFormula.replace(/\/>$/, `>${formulaXml}</${prefix}c>`);
  return withoutFormula.replace(/^(<(?:[A-Za-z_][\w.-]*:)?c\b[^>]*>)/, (openingTag) => `${openingTag}${formulaXml}`);
}

function escapeXml(value) { return String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;"); }

function updateSummarySheet(workbook, history, holdings, requestedDate, bootstrap) {
  const sheet = workbook.worksheets.getItem("Summary");
  const cashflowColumns = detectCashflowColumns(workbook.worksheets.getItem("CashFlow"));
  const oldLayout = detectLayout(sheet);
  const oldRows = readExistingHistory(sheet, oldLayout);
  const oldAnchor = readOpeningAnchor(sheet, oldLayout);
  const openingDate = excelDateText(oldAnchor.date);
  const oldTickers = oldLayout.tickers;
  const names = new Map(holdings.map((item) => [item.ticker, item.name]));
  const displayTickers = new Map(oldLayout.displayTickers.map((ticker, index) => [oldTickers[index], ticker]));
  for (const item of holdings) if (!displayTickers.has(item.ticker)) displayTickers.set(item.ticker, item.stockTicker);
  const tickers = [...oldTickers];
  for (const item of holdings) if (!tickers.includes(item.ticker)) tickers.push(item.ticker);
  for (const ticker of tickers) if (!names.has(ticker)) names.set(ticker, ticker);

  const incoming = history.filter((item) => item.date >= openingDate && item.date <= requestedDate);
  if (!incoming.length) fail(`Price CSV has no common trading-day prices through ${requestedDate}.`);
  for (const item of incoming) {
    for (const ticker of tickers) {
      const quote = item.prices.get(ticker);
      if (!quote || !Number.isFinite(quote.price) || quote.price <= 0) fail(`Missing valid USD close for ${ticker} on ${item.date}.`);
      if (quote.currency !== "USD") fail(`Summary currently expects USD prices; ${ticker} is ${quote.currency} on ${item.date}.`);
    }
  }

  const rowsByDate = new Map(oldRows.map((item) => [item.date, item]));
  let addedRows = 0;
  let updatedRows = 0;
  for (const item of incoming) {
    const existing = rowsByDate.get(item.date);
    if (existing) updatedRows += 1;
    else addedRows += 1;
    const prices = new Map(existing?.prices || []);
    for (const [ticker, quote] of item.prices) prices.set(ticker, quote.price);
    rowsByDate.set(item.date, { date: item.date, dateValue: item.dateValue, prices });
  }
  const rows = [...rowsByDate.values()].sort((a, b) => a.date.localeCompare(b.date));
  if (rows.length > MAX_DATA_ROW - FIRST_DATA_ROW + 1) fail(`Summary history exceeds row ${MAX_DATA_ROW}.`);

  const layout = makeLayout(tickers.length);
  const oldTable = sheet.tables.items.find((item) => ["SummaryHistory", "Table1"].includes(item.name));
  const oldTableEnd = oldTable ? rangeEndRow(oldTable.getRange().address) : FIRST_DATA_ROW + rows.length + 20;
  const tableEnd = Math.max(oldTableEnd || 0, FIRST_DATA_ROW + rows.length + 20);
  if (oldTable) oldTable.delete();

  const clearEndCol = Math.max(oldLayout.anchorValueCol, layout.anchorValueCol) + 2;
  sheet.getRangeByIndexes(9, 0, MAX_DATA_ROW - 8, clearEndCol).clear({ applyTo: "contents" });
  if (layout.priceEndCol > oldLayout.priceEndCol) {
    // These columns used to be calculated metrics but are becoming raw ticker
    // prices. Clear formulas and formats together so imported table metadata
    // cannot restore the former calculated-column formulas on export.
    sheet.getRangeByIndexes(
      9,
      oldLayout.priceEndCol + 1,
      MAX_DATA_ROW - 8,
      layout.priceEndCol - oldLayout.priceEndCol,
    ).clear({ applyTo: "all" });
  }
  writeHeaders(sheet, layout, tickers, names, displayTickers, bootstrap);
  writeHistory(sheet, layout, tickers, rows, cashflowColumns, bootstrap);
  writeOpeningAnchor(sheet, layout, oldAnchor);
  writeKpis(sheet, layout);
  formatSummary(sheet, layout, rows.length);

  const table = sheet.tables.add(`A13:${columnLetter(layout.totalNetWorthCol)}${tableEnd}`, true, "Table1");
  table.style = "TableStyleMedium2";
  // Re-creating an imported Excel table can restore isolated calculated-column
  // formulas from the former metric column. Reassert that Date and ticker-price
  // columns contain only typed values after the table exists.
  writeDateAndPriceValues(sheet, tickers, rows);
  updateChart(sheet, layout, rows.length);

  return {
    tickers,
    displayTickers: Object.fromEntries(tickers.map((ticker) => [ticker, displayTickers.get(ticker)])),
    addedTickers: tickers.filter((ticker) => !oldTickers.includes(ticker)),
    names: Object.fromEntries(tickers.map((ticker) => [ticker, names.get(ticker)])),
    tradingDaysInRequestedMonth: incoming.length,
    tradingDaysInRequestedRange: incoming.length,
    historyStartDate: openingDate,
    addedRows,
    updatedRows,
    totalHistoryRows: rows.length,
    latestPriceDate: rows.at(-1)?.date || null,
    cashflowBaseColumns: cashflowColumns,
    layout: Object.fromEntries(Object.entries(layout).map(([key, value]) => [key, typeof value === "number" ? columnLetter(value) : value])),
  };
}

function detectLayout(sheet) {
  const headers = sheet.getRange("A13:AZ13").values[0].map(clean);
  const metricStartCol = headers.indexOf(SUMMARY_HEADERS[0]);
  const cashCol = headers.indexOf("实时持有现金");
  if (metricStartCol < 1 || cashCol < 0) fail("Could not detect the current Summary layout.");
  const displayTickers = headers.slice(1, metricStartCol).map(clean).filter(Boolean);
  const tickers = displayTickers.map(normalizeTicker);
  const anchorValueCol = findOpeningAnchorValueCol(sheet);
  return {
    tickers,
    displayTickers,
    priceStartCol: 1,
    priceEndCol: metricStartCol - 1,
    metricStartCol,
    totalNetWorthCol: metricStartCol + 2,
    helperStartCol: cashCol - tickers.length,
    helperEndCol: cashCol - 1,
    cashCol,
    anchorValueCol,
  };
}

function detectCashflowColumns(sheet) {
  const headers = sheet.getRange("A11:AN11").values[0].map(clean);
  const convertedHeaders = new Set(["Amount (Base)", "Converted Amount"]);
  const incomeIndex = headers.findIndex((header, index) => index < 13 && convertedHeaders.has(header));
  const expenseIndex = headers.findIndex((header, index) => index >= 13 && index < 26 && convertedHeaders.has(header));
  if (incomeIndex < 0 || expenseIndex < 0) fail("Could not detect the Income and Expense base-amount columns in CashFlow row 11.");
  return { incomeBase: columnLetter(incomeIndex), expenseBase: columnLetter(expenseIndex) };
}

function makeLayout(tickerCount) {
  const metricStartCol = 1 + tickerCount;
  const helperStartCol = metricStartCol + 4;
  const helperEndCol = helperStartCol + tickerCount - 1;
  const cashCol = helperEndCol + 1;
  return {
    priceStartCol: 1,
    priceEndCol: tickerCount,
    metricStartCol,
    stockMvUsdCol: metricStartCol,
    stockMvBaseCol: metricStartCol + 1,
    totalNetWorthCol: metricStartCol + 2,
    separatorCol: metricStartCol + 3,
    helperStartCol,
    helperEndCol,
    cashCol,
    anchorLabelCol: cashCol,
    anchorValueCol: cashCol + 1,
  };
}

function readExistingHistory(sheet, layout) {
  const values = sheet.getRangeByIndexes(FIRST_DATA_ROW - 1, 0, MAX_DATA_ROW - FIRST_DATA_ROW + 1, layout.priceEndCol + 1).values;
  const rows = [];
  for (const row of values) {
    const date = excelDateText(row[0]);
    if (!date) continue;
    const prices = new Map();
    layout.tickers.forEach((ticker, index) => {
      const price = Number(row[index + 1]);
      if (Number.isFinite(price) && price > 0) prices.set(ticker, price);
    });
    rows.push({ date, dateValue: new Date(`${date}T00:00:00.000Z`), prices });
  }
  return rows;
}

function readStockHoldings(sheet) {
  const values = sheet.getRange(`P${HOLDINGS_FIRST_ROW}:Q${HOLDINGS_LAST_ROW}`).values;
  const holdings = [];
  const seen = new Set();
  for (const [rawTicker, rawName] of values) {
    const ticker = normalizeTicker(rawTicker);
    if (!ticker || ticker === "TICKER" || seen.has(ticker)) continue;
    seen.add(ticker);
    holdings.push({ ticker, stockTicker: clean(rawTicker), name: canonicalSecurityName(ticker, rawName) });
  }
  // Closed positions are absent from the current-holdings block but still
  // matter to historical net worth before their sale date.
  for (const [, , rawTicker, rawName] of sheet.getRange(`A${STOCK_FIRST_ROW}:D${STOCK_LAST_ROW}`).values) {
    const ticker = normalizeTicker(rawTicker);
    if (!ticker || ticker === "TICKER" || seen.has(ticker)) continue;
    seen.add(ticker);
    holdings.push({ ticker, stockTicker: clean(rawTicker), name: canonicalSecurityName(ticker, rawName) });
  }
  return holdings;
}

function writeHeaders(sheet, layout, tickers, names, displayTickers, bootstrap) {
  const displayed = tickers.map((ticker) => displayTickers.get(ticker) || ticker);
  sheet.getRangeByIndexes(9, 0, 1, tickers.length + 1).values = [["Ticker", ...displayed]];
  sheet.getRangeByIndexes(10, 0, 1, tickers.length + 1).values = [["Units Held", ...Array(tickers.length).fill(null)]];
  // The reference Summary keeps row 12 blank in the Linked Holdings/price
  // area. Security names remain available in Stock and are not repeated here.
  sheet.getRangeByIndexes(11, 0, 1, tickers.length + 1).values = [Array(tickers.length + 1).fill(null)];
  for (let index = 0; index < tickers.length; index += 1) {
    const col = columnLetter(index + 1);
    const opening = bootstrap.openingHoldings.find((item) => item.ticker === tickers[index]) || null;
    sheet.getRange(`${col}11`).formulas = [[unitsHeldFormula(`${col}$10`, opening, bootstrap.openingDateText)]];
  }
  const mainHeaders = ["Date", ...displayed, ...SUMMARY_HEADERS];
  sheet.getRangeByIndexes(12, 0, 1, mainHeaders.length).values = [mainHeaders];
  sheet.getRangeByIndexes(11, layout.helperStartCol, 1, 1).values = [["实时持有数量"]];
  sheet.getRangeByIndexes(12, layout.helperStartCol, 1, tickers.length + 1).values = [[...displayed.map((ticker) => `${ticker} Units`), "实时持有现金"]];
}

function writeHistory(sheet, layout, tickers, rows, cashflowColumns, bootstrap) {
  if (!rows.length) return;
  writeDateAndPriceValues(sheet, tickers, rows);
  for (let index = 0; index < rows.length; index += 1) {
    const row = FIRST_DATA_ROW + index;
    const priceStart = columnLetter(layout.priceStartCol);
    const priceEnd = columnLetter(layout.priceEndCol);
    const helperStart = columnLetter(layout.helperStartCol);
    const helperEnd = columnLetter(layout.helperEndCol);
    const stockUsd = columnLetter(layout.stockMvUsdCol);
    const stockBase = columnLetter(layout.stockMvBaseCol);
    const total = columnLetter(layout.totalNetWorthCol);
    const cash = columnLetter(layout.cashCol);
    const anchor = columnLetter(layout.anchorValueCol);
    sheet.getRangeByIndexes(row - 1, layout.metricStartCol, 1, 2).formulas = [[
      `=IF(COUNTA(${priceStart}${row}:${priceEnd}${row})=0,"",SUMPRODUCT(${priceStart}${row}:${priceEnd}${row},${helperStart}${row}:${helperEnd}${row}))`,
      `=${stockUsd}${row}*IF('Asset'!$K$3="USD",1,'Asset'!$K$4)`,
    ]];
    for (let tickerIndex = 0; tickerIndex < tickers.length; tickerIndex += 1) {
      const helperCol = columnLetter(layout.helperStartCol + tickerIndex);
      const opening = bootstrap.openingHoldings.find((item) => item.ticker === tickers[tickerIndex]) || null;
      sheet.getRange(`${helperCol}${row}`).formulas = [[historicalUnitsFormula(`${helperCol}$13`, row, opening, bootstrap.openingDateText)]];
    }
    sheet.getRange(`${cash}${row}`).formulas = [[liveCashFormula(row, anchor, cashflowColumns)]];
    // Write the dependent total after both Stock MV and live cash are present so
    // the calculation cache does not retain a transient #N/A on export.
    sheet.getRange(`${total}${row}`).formulas = [[`=${stockBase}${row}+${cash}${row}`]];
  }
}

function writeDateAndPriceValues(sheet, tickers, rows) {
  const mainValues = rows.map((item) => [item.dateValue, ...tickers.map((ticker) => item.prices.get(ticker) ?? null)]);
  const dateAndPriceRange = sheet.getRangeByIndexes(FIRST_DATA_ROW - 1, 0, rows.length, tickers.length + 1);
  // A new ticker can occupy a column that previously contained a calculated
  // metric. Clear the old formula matrix explicitly before writing typed
  // dates/prices; assigning null values alone can preserve a table formula.
  dateAndPriceRange.clear({ applyTo: "contents" });
  dateAndPriceRange.formulas = rows.map(() => Array(tickers.length + 1).fill(""));
  dateAndPriceRange.values = mainValues;
}

function unitsHeldFormula(tickerCell, opening, openingDate) {
  const baseline = opening?.units || 0;
  const date = excelDateFormula(openingDate);
  return `=${baseline}+SUMIFS('Stock'!$F$${STOCK_FIRST_ROW}:$F$${STOCK_LAST_ROW},'Stock'!$C$${STOCK_FIRST_ROW}:$C$${STOCK_LAST_ROW},${tickerCell},'Stock'!$E$${STOCK_FIRST_ROW}:$E$${STOCK_LAST_ROW},"Buy",'Stock'!$B$${STOCK_FIRST_ROW}:$B$${STOCK_LAST_ROW},">="&${date})-SUMIFS('Stock'!$F$${STOCK_FIRST_ROW}:$F$${STOCK_LAST_ROW},'Stock'!$C$${STOCK_FIRST_ROW}:$C$${STOCK_LAST_ROW},${tickerCell},'Stock'!$E$${STOCK_FIRST_ROW}:$E$${STOCK_LAST_ROW},"Sell",'Stock'!$B$${STOCK_FIRST_ROW}:$B$${STOCK_LAST_ROW},">="&${date})`;
}

function historicalUnitsFormula(headerCell, row, opening, openingDate) {
  const baseline = opening?.units || 0;
  const date = excelDateFormula(openingDate);
  return `=${baseline}+SUMIFS('Stock'!$F$${STOCK_FIRST_ROW}:$F$${STOCK_LAST_ROW},'Stock'!$C$${STOCK_FIRST_ROW}:$C$${STOCK_LAST_ROW},LEFT(${headerCell},FIND(" ",${headerCell})-1),'Stock'!$E$${STOCK_FIRST_ROW}:$E$${STOCK_LAST_ROW},"Buy",'Stock'!$B$${STOCK_FIRST_ROW}:$B$${STOCK_LAST_ROW},">="&${date},'Stock'!$B$${STOCK_FIRST_ROW}:$B$${STOCK_LAST_ROW},"<="&$A${row})-SUMIFS('Stock'!$F$${STOCK_FIRST_ROW}:$F$${STOCK_LAST_ROW},'Stock'!$C$${STOCK_FIRST_ROW}:$C$${STOCK_LAST_ROW},LEFT(${headerCell},FIND(" ",${headerCell})-1),'Stock'!$E$${STOCK_FIRST_ROW}:$E$${STOCK_LAST_ROW},"Sell",'Stock'!$B$${STOCK_FIRST_ROW}:$B$${STOCK_LAST_ROW},">="&${date},'Stock'!$B$${STOCK_FIRST_ROW}:$B$${STOCK_LAST_ROW},"<="&$A${row})`;
}

function excelDateFormula(date) {
  return `DATE(${date.slice(0, 4)},${Number(date.slice(5, 7))},${Number(date.slice(8, 10))})`;
}

function liveCashFormula(row, anchorCol, cashflowColumns) {
  const income = cashflowColumns.incomeBase;
  const expense = cashflowColumns.expenseBase;
  return `=IF(A${row}="",NA(),$${anchorCol}$8+SUMIFS('CashFlow'!$${income}$12:$${income}$1000,'CashFlow'!$A$12:$A$1000,">="&$${anchorCol}$7,'CashFlow'!$A$12:$A$1000,"<="&A${row})-SUMIFS('CashFlow'!$${expense}$12:$${expense}$1000,'CashFlow'!$N$12:$N$1000,">="&$${anchorCol}$7,'CashFlow'!$N$12:$N$1000,"<="&A${row})-SUMIFS('Stock'!$H$21:$H$1004,'Stock'!$B$21:$B$1004,">="&$${anchorCol}$7,'Stock'!$B$21:$B$1004,"<="&A${row},'Stock'!$E$21:$E$1004,"Buy")+SUMIFS('Stock'!$H$21:$H$1004,'Stock'!$B$21:$B$1004,">="&$${anchorCol}$7,'Stock'!$B$21:$B$1004,"<="&A${row},'Stock'!$E$21:$E$1004,"Sell")+SUMIFS('Stock'!$O$21:$O$1004,'Stock'!$B$21:$B$1004,">="&$${anchorCol}$7,'Stock'!$B$21:$B$1004,"<="&A${row},'Stock'!$E$21:$E$1004,"Dividend"))`;
}

function readOpeningAnchor(sheet, layout) {
  const values = sheet.getRangeByIndexes(6, layout.anchorValueCol, 2, 1).values.flat();
  if (!excelDateText(values[0]) || !Number.isFinite(Number(values[1]))) fail("Summary opening date or opening cash balance is missing.");
  return { date: values[0], cash: Number(values[1]) };
}

function writeOpeningAnchor(sheet, layout, anchor) {
  sheet.getRangeByIndexes(6, layout.anchorLabelCol, 2, 1).values = [["Opening Date"], ["Opening Cash Balance (Base)"]];
  sheet.getRangeByIndexes(6, layout.anchorValueCol, 2, 1).values = [[anchor.date], [anchor.cash]];
  sheet.getRangeByIndexes(6, layout.anchorValueCol, 1, 1).format.numberFormat = "yyyy-mm-dd";
  sheet.getRangeByIndexes(7, layout.anchorValueCol, 1, 1).format.numberFormat = "#,##0.00;[Red](#,##0.00)";
}

function writeKpis(sheet, layout) {
  const stockBase = columnLetter(layout.stockMvBaseCol);
  const total = columnLetter(layout.totalNetWorthCol);
  sheet.getRange("B4").formulas = [[`=IFERROR(INDEX($${total}$14:$${total}$1046,MATCH(MAX($A$14:$A$1046),$A$14:$A$1046,0)),0)`]];
  sheet.getRange("B5").formulas = [["=IFERROR(MAX($A$14:$A$1046),\"\")"]];
  sheet.getRange("B6").formulas = [[`=IFERROR(INDEX($${stockBase}$14:$${stockBase}$1046,MATCH(MAX($A$14:$A$1046),$A$14:$A$1046,0)),0)`]];
  writeSummaryReconciliation(sheet, total);
}

export function writeSummaryReconciliation(sheet, totalNetWorthColumn = null) {
  const total = totalNetWorthColumn || detectSummaryHeaderColumn(sheet, "Total Net Worth (Base)");
  sheet.getRange("N4:N7").values = [["Real Asset"], ["Calculated Asset"], ["Total Difference"], ["Status"]];
  sheet.getRange("O4").formulas = [["='Asset'!B9"]];
  sheet.getRange("O5").formulas = [[`=IFERROR(INDEX($${total}$14:$${total}$1046,MATCH(MAX($A$14:$A$1046),$A$14:$A$1046,0)),0)`]];
  sheet.getRange("O6").formulas = [["=O4-O5"]];
  sheet.getRange("O7").formulas = [["=IF(ABS(O6)<100,\"OK\",\"Check\")"]];
}

function detectSummaryHeaderColumn(sheet, header) {
  const headers = sheet.getRange("A13:AZ13").values[0].map(clean);
  const index = headers.indexOf(header);
  if (index < 0) fail(`Could not find Summary header: ${header}`);
  return columnLetter(index);
}

function formatSummary(sheet, layout, rowCount) {
  const lastRow = Math.max(FIRST_DATA_ROW, FIRST_DATA_ROW + rowCount - 1);
  const lastCol = columnLetter(layout.anchorValueCol);
  sheet.getRange(`A10:${lastCol}${lastRow}`).format.font = { name: "Calibri", size: 11 };
  sheet.getRange(`A10:${lastCol}${lastRow}`).format.horizontalAlignment = "center";
  sheet.getRangeByIndexes(11, 0, 1, layout.priceEndCol + 1).format.wrapText = false;
  sheet.getRange(`A14:A${lastRow}`).format.numberFormat = "yyyy-mm-dd";
  sheet.getRange(`${columnLetter(layout.priceStartCol)}14:${columnLetter(layout.totalNetWorthCol)}${lastRow}`).format.numberFormat = "#,##0.00;[Red](#,##0.00)";
  sheet.getRange(`${columnLetter(layout.helperStartCol)}14:${columnLetter(layout.helperEndCol)}${lastRow}`).format.numberFormat = "0.0000";
  sheet.getRange(`${columnLetter(layout.cashCol)}14:${columnLetter(layout.cashCol)}${lastRow}`).format.numberFormat = "#,##0.00;[Red](#,##0.00)";
  sheet.getRange("A:A").format.columnWidth = 12;
  sheet.getRange(`${columnLetter(layout.priceStartCol)}:${columnLetter(layout.priceEndCol)}`).format.columnWidth = 12;
  sheet.getRange(`${columnLetter(layout.metricStartCol)}:${columnLetter(layout.totalNetWorthCol)}`).format.columnWidth = 21;
  sheet.getRange(`${columnLetter(layout.helperStartCol)}:${columnLetter(layout.cashCol)}`).format.columnWidth = 14;
}

function updateChart(sheet, layout, rowCount) {
  const charts = sheet.charts.items;
  if (charts.length !== 1 || charts[0].series.items.length !== 1) fail("Unexpected Summary chart structure; expected one chart with one series.");
  const lastRow = FIRST_DATA_ROW + rowCount - 1;
  const total = columnLetter(layout.totalNetWorthCol);
  charts[0].series.items[0].categoryFormula = `Summary!$A$14:$A$${lastRow}`;
  charts[0].series.items[0].formula = `Summary!$${total}$14:$${total}$${lastRow}`;
}

async function fetchMonthlyPrices(inputDir, tickers, summaryDate, startDate) {
  const python = process.env.WEALTH_PYTHON;
  if (!python) fail("WEALTH_PYTHON is required when --price-csv is not supplied.");
  const temporaryDirectory = await fs.mkdtemp(path.join(os.tmpdir(), "wealth-summary-prices-"));
  const output = path.join(temporaryDirectory, "prices.csv");
  const args = [DEFAULT_PRICE_SCRIPT, "--output", output, "--tickers", tickers.join(",")];
  if (startDate) args.push("--start", startDate);
  if (summaryDate) args.push("--as-of", summaryDate);
  else {
    const discovery = await discoverCashflowCsvs(path.join(path.resolve(inputDir), "CashFlow"));
    for (const file of discovery.recognized) args.push("--cashflow-csv", file);
  }
  const result = await execFileAsync(python, args, { maxBuffer: 10 * 1024 * 1024 });
  const lines = result.stdout.trim().split(/\r?\n/).filter(Boolean);
  const fetchReport = JSON.parse(lines.at(-1));
  return { file: output, fetched: true, fetchReport };
}

async function readPriceHistory(csvPath) {
  const text = (await fs.readFile(csvPath, "utf8")).replace(/^\uFEFF/, "");
  const workbook = await Workbook.fromCSV(text, { sheetName: "Prices" });
  const values = workbook.worksheets.getItem("Prices").getUsedRange(true).values;
  if (!values?.length) fail(`Price CSV is empty: ${csvPath}`);
  const headers = values[0].map(clean);
  const tickerCol = findHeader(headers, ["Ticker", "Symbol"]);
  const priceCol = findHeader(headers, ["Current Price", "Close", "Price"]);
  const currencyCol = findHeader(headers, ["Currency", "币种"]);
  const dateCol = findHeader(headers, ["Date", "更新日期"]);
  const grouped = new Map();
  for (const row of values.slice(1)) {
    const ticker = normalizeTicker(row[tickerCol]);
    const date = excelDateText(row[dateCol]);
    const price = Number(row[priceCol]);
    const currency = clean(row[currencyCol]).toUpperCase();
    if (!ticker || !date || !Number.isFinite(price)) continue;
    if (!grouped.has(date)) grouped.set(date, new Map());
    grouped.get(date).set(ticker, { price, currency });
  }
  return [...grouped.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([date, prices]) => ({ date, dateValue: new Date(`${date}T00:00:00.000Z`), prices }));
}

function resolveRequestedDate(explicitDate, fetchReport, history) {
  const date = explicitDate || fetchReport?.requested_as_of || history.at(-1)?.date;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date || "")) fail(`Invalid or missing Summary date: ${date || "blank"}`);
  return date;
}

function findOpeningAnchorValueCol(sheet) {
  const labels = sheet.getRange("A7:AZ7").values[0].map(clean);
  const labelCol = labels.indexOf("Opening Date");
  if (labelCol < 0) fail("Could not find Summary opening-date anchor.");
  return labelCol + 1;
}

function findHeader(headers, choices) {
  const index = headers.findIndex((header) => choices.includes(header));
  if (index < 0) fail(`Price CSV is missing one of these headers: ${choices.join(", ")}`);
  return index;
}

function rangeEndRow(address) { return Number(String(address).split(":").at(-1).match(/\d+$/)?.[0]); }
function normalizeTicker(value) { const ticker = clean(value).toUpperCase(); return ticker === "ECHO" ? "SATS" : ticker; }
function excelDateText(value) { if (value instanceof Date) return value.toISOString().slice(0, 10); if (typeof value === "number" && Number.isFinite(value)) return new Date(Date.UTC(1899, 11, 30) + Math.round(value) * 86400000).toISOString().slice(0, 10); return clean(value).slice(0, 10); }
function clean(value) { return value === null || value === undefined ? "" : String(value).trim(); }
function parseCellAddress(address) { const match = String(address).match(/^([A-Z]+)(\d+)$/); if (!match) fail(`Invalid cell address: ${address}`); return { col: columnIndex(match[1]), row: Number(match[2]) }; }
function columnIndex(column) { let result = 0; for (const character of String(column)) result = result * 26 + character.charCodeAt(0) - 64; return result - 1; }
function columnLetter(index) { let value = index + 1; let result = ""; while (value > 0) { const remainder = (value - 1) % 26; result = String.fromCharCode(65 + remainder) + result; value = Math.floor((value - 1) / 26); } return result; }
async function exists(file) { try { await fs.access(file); return true; } catch { return false; } }
function fail(message) { throw new Error(message); }

function parseCliArgs(argv) {
  const args = {};
  for (let index = 0; index < argv.length; index += 1) {
    const key = argv[index];
    if (key === "--dry-run") args.dryRun = true;
    else if (key === "--allow-existing-output") args.allowExistingOutput = true;
    else if (key.startsWith("--")) args[key.slice(2).replace(/-([a-z])/g, (_, character) => character.toUpperCase())] = argv[++index];
    else fail(`Unexpected argument: ${key}`);
  }
  return args;
}

async function main() {
  const args = parseCliArgs(process.argv.slice(2));
  const inputDir = path.resolve(args.inputDir || process.cwd());
  const workbookPath = path.resolve(args.workbook || path.join(inputDir, "Output", "One_Piece.xlsx"));
  if (!args.output) fail("Usage: update_summary.mjs [--input-dir DIR] [--input-file FILE] [--workbook FILE] --output FILE [--summary-date YYYY-MM-DD] [--price-csv FILE] [--dry-run] [--allow-existing-output]");
  const report = await runSummaryUpdate({
    workbookPath,
    outputPath: args.output,
    inputDir,
    inputFile: args.inputFile || path.join(inputDir, "monthly-close-input.md"),
    priceCsv: args.priceCsv || null,
    summaryDate: args.summaryDate || null,
    dryRun: Boolean(args.dryRun),
    allowExistingOutput: Boolean(args.allowExistingOutput),
  });
  console.log(JSON.stringify(report, null, 2));
}

const directRun = process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href;
if (directRun) main().catch((error) => { console.error(error instanceof Error ? error.message : String(error)); process.exitCode = 1; });

#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import { createRequire } from "node:module";
import { pathToFileURL } from "node:url";
import { saveWorkbookFormulaSafe } from "./workbook_io.mjs";

const dependencies = process.env.CODEX_NODE_MODULES;
if (!dependencies) throw new Error("CODEX_NODE_MODULES is required");
const require = createRequire(path.join(dependencies, "package.json"));
const { FileBlob, SpreadsheetFile } = require("@oai/artifact-tool");

const BLUE = "#1F4E78";
const LIGHT_BLUE = "#D9EAF7";
const PALE_BLUE = "#EAF3FF";
const GREEN = "#9BBB59";
const PURPLE = "#EFE5FF";
const WHITE = "#FFFFFF";
const BLACK = "#000000";
const DARK_TEXT = "#1F1F1F";
const RED = "#FF0000";
const INPUT_BLUE = "#0000FF";
const LINK_GREEN = "#008000";
const YELLOW = "#FFF2CC";
const GOLD = "#FFC000";
const GRID = "#BFBFBF";
const CJK = "Noto Sans CJK SC";
const VALUE = "Cambria";
const BODY = "Calibri";
const MONEY = "#,##0.00;[Red](#,##0.00);-";
const UNITS = "0.0000;[Red](0.0000);-";

export function applyWorkbookFormat(workbook) {
  formatCashflow(workbook.worksheets.getItem("CashFlow"));
  formatStock(workbook.worksheets.getItem("Stock"));
  formatAsset(workbook.worksheets.getItem("Asset"));
  formatSummary(workbook.worksheets.getItem("Summary"));
  return {
    sheets: ["CashFlow", "Stock", "Asset", "Summary"],
    baseline: "Output/bootstrap/One_Piece.xlsx",
    runtimeTemplateDependency: false,
  };
}

export async function runWorkbookFormat({ workbookPath, outputPath, allowExistingOutput = false }) {
  const source = path.resolve(workbookPath);
  const destination = path.resolve(outputPath);
  if (source === destination) fail("Refusing to overwrite the source workbook. Choose a distinct output path.");
  if (!allowExistingOutput && await exists(destination)) fail(`Output already exists: ${destination}`);
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(source));
  const report = applyWorkbookFormat(workbook);
  await fs.mkdir(path.dirname(destination), { recursive: true });
  await saveWorkbookFormulaSafe(workbook, destination);
  return { workbook: source, output: destination, ...report };
}

function formatCashflow(sheet) {
  sheet.showGridLines = false;
  const cashflowCanvas = sheet.getRange("A1:AN500");
  cashflowCanvas.format.font = { name: BODY, size: 10, color: BLACK };
  cashflowCanvas.format.horizontalAlignment = "center";
  cashflowCanvas.format.verticalAlignment = "center";
  cashflowCanvas.format.wrapText = false;
  fixedMergedBand(sheet, "A1:L1", "Summary", { font: { name: CJK, size: 12, bold: true, color: WHITE } });
  for (const [address, title] of [["A10:L10", "Income"], ["N10:Y10", "Expenses"], ["AA10:AN10", "Transfer"]]) {
    fixedMergedBand(sheet, address, title, { font: { name: CJK, size: 12, bold: true, color: WHITE } });
  }
  for (const address of ["A11:J11", "N11:W11", "AA11:AM11"]) {
    band(sheet.getRange(address), CJK, 11);
    bottomBorder(sheet.getRange(address));
  }
  sheet.getRange("A2").values = [["Selected Month："]];
  for (const address of ["A2:A5", "D3:D5", "G3:G5", "J3:J5"]) {
    sheet.getRange(address).format.font = { name: CJK, size: 11, bold: true, color: DARK_TEXT };
  }
  sheet.getRange("B2").format.fill = LIGHT_BLUE;
  sheet.getRange("B2").format.font = { name: VALUE, size: 11, bold: true, color: INPUT_BLUE };
  sheet.getRange("B2").format.horizontalAlignment = "center";
  for (const address of ["E3:E5", "H3:H5", "K3:K5"]) sheet.getRange(address).format.fill = LIGHT_BLUE;
  for (const address of ["B3:B5", "E3:E5", "H3:H5", "K3:K5"]) {
    sheet.getRange(address).format.font = { name: VALUE, size: 11, bold: true, color: BLACK };
  }
  sheet.getRange("A7").format = { fill: GOLD, font: { name: BODY, size: 14, bold: true, color: BLACK }, verticalAlignment: "center" };
  sheet.getRange("B7").format.fill = WHITE;
  sheet.getRange("B7").format.font = { name: BODY, size: 16, bold: true, color: BLACK };
  sheet.getRange("B3:B5").format.numberFormat = MONEY;
  sheet.getRange("B7").format.numberFormat = MONEY;
  sheet.getRange("E3:E5").format.numberFormat = "yyyy-mm-dd";
  sheet.getRange("H3:H5").format.numberFormat = "yyyy-mm-dd";
  grid(sheet.getRange("A2:B5"));
  grid(sheet.getRange("D3:E5"));
  grid(sheet.getRange("G3:H5"));
  grid(sheet.getRange("J3:K5"));
  grid(sheet.getRange("A7:B7"));
  for (const address of ["A12:J454", "N12:W456", "AA12:AM500"]) {
    const range = sheet.getRange(address);
    range.format.font = { name: BODY, size: 10, color: LINK_GREEN };
    range.format.horizontalAlignment = "center";
    grid(range);
  }
  for (const address of ["C12:E454", "H12:H454", "P12:R456", "U12:U456", "AC12:AF500", "AI12:AJ500", "AM12:AM500"]) {
    sheet.getRange(address).format.font = { name: CJK, size: 10, color: LINK_GREEN };
  }
  for (const address of ["J12:J454", "W12:W456", "AK12:AL500"]) {
    sheet.getRange(address).format.font = { name: BODY, size: 11, color: BLACK };
  }
  for (const address of ["A12:A454", "N12:N456", "AA12:AA500"]) sheet.getRange(address).format.numberFormat = "yyyy-mm-dd";
  for (const address of ["F12:F454", "J12:J454", "S12:S456", "W12:W456", "AG12:AG500", "AK12:AL500"]) sheet.getRange(address).format.numberFormat = MONEY;
  setColumnWidths(sheet, {
    A: 27.83, B: 13.83, C: 12.66, D: 11.83, E: 20, F: 11.5, G: 12.66, H: 13, I: 8, J: 18.16,
    N: 23.16, O: 11.83, P: 9.33, Q: 14.83, R: 23.66, S: 13, T: 13, U: 13, V: 8.16, W: 18.16,
    AA: 13, AB: 13, AC: 13, AD: 12.33, AE: 13, AF: 15.66, AG: 13, AH: 13, AI: 13, AJ: 13, AK: 8, AL: 18.33, AM: 13.66,
  });
  sheet.getRange("A1:AM1").format.rowHeight = 16;
  sheet.getRange("A7:AM7").format.rowHeight = 21;
  for (const address of ["A12:AM30", "A10:AM10"]) sheet.getRange(address).format.rowHeight = address.endsWith("10") ? 16 : 15;
}

function formatStock(sheet) {
  sheet.showGridLines = false;
  const stockCanvas = sheet.getRange("A1:AI263");
  stockCanvas.format.font = { name: VALUE, size: 11, color: BLACK };
  stockCanvas.format.verticalAlignment = "center";
  stockCanvas.format.wrapText = false;
  fixedMergedTitle(sheet, "A1:O1", "Stock buying, selling, and holding analysis", { name: CJK, size: 14, bold: true, color: BLACK });
  unmergeIfPresent(sheet, "P2:S2");
  sheet.getRange("P2:S2").format.fill = WHITE;
  sheet.getRange("P2:S2").format.font = { name: CJK, size: 11, bold: true, color: BLACK };
  sheet.getRange("P2:S2").format.horizontalAlignment = "left";
  sheet.getRange("P2").values = [["当前价格输入（可手动更新）"]];
  band(sheet.getRange("P3:S3"), VALUE, 11);
  fixedMergedTitle(sheet, "A19:O19", "交易明细", { name: CJK, size: 16, bold: true, color: BLACK }, LIGHT_BLUE);
  band(sheet.getRange("A20:O20"), VALUE, 11);
  unmergeIfPresent(sheet, "P36:Z36");
  sheet.getRange("P36:Z36").format.fill = GREEN;
  sheet.getRange("P36:Z36").format.font = { name: CJK, size: 16, bold: true, color: BLACK };
  sheet.getRange("P36:Z36").format.horizontalAlignment = "left";
  sheet.getRange("P36").values = [["持仓汇总（按当前市值）"]];
  band(sheet.getRange("P37:Z37"), VALUE, 11);
  colorBand(sheet.getRange("A9:A10"), BLUE, WHITE, CJK, 11);
  sheet.getRange("A3:A6").format.font = { name: CJK, size: 11, bold: true, color: BLACK };
  sheet.getRange("B3:B6").format.font = { name: VALUE, size: 11, color: BLACK };
  sheet.getRange("B9").format.font = { name: VALUE, size: 11, color: LINK_GREEN };
  sheet.getRange("B10").format.font = { name: VALUE, size: 11, color: INPUT_BLUE };
  sheet.getRange("B3:B5").format.numberFormat = MONEY;
  sheet.getRange("B6").format.numberFormat = "#,##0";
  sheet.getRange("A21:O212").format.fill = WHITE;
  sheet.getRange("A21:O212").format.font = { name: VALUE, size: 11, color: BLACK };
  sheet.getRange("P4:S33").format.fill = WHITE;
  sheet.getRange("P4:S33").format.font = { name: VALUE, size: 11, color: BLACK };
  sheet.getRange("P38:Z67").format.fill = WHITE;
  sheet.getRange("P38:Z67").format.font = { name: VALUE, size: 11, color: BLACK };
  // AA21 is outside every current Stock block. Remove the dark-blue style
  // retained by the older Bootstrap holdings header.
  sheet.getRange("AA21").clear({ applyTo: "formats" });
  const lastTransactionRow = lastPopulatedRow(sheet.getRange("A21:A212").values, 21);
  const lastPriceRow = lastPopulatedRow(sheet.getRange("P4:P33").values, 4);
  const lastHoldingRow = lastPopulatedRow(sheet.getRange("P38:P67").values, 38);
  if (lastTransactionRow >= 21) {
    sheet.getRange(`A21:O${lastTransactionRow}`).format.fill = WHITE;
    grid(sheet.getRange(`A20:O${lastTransactionRow}`));
    sheet.getRange(`A21:A${lastTransactionRow}`).format.horizontalAlignment = "left";
    sheet.getRange(`B21:B${lastTransactionRow}`).format.horizontalAlignment = "center";
    sheet.getRange(`C21:D${lastTransactionRow}`).format.horizontalAlignment = "left";
    sheet.getRange(`E21:O${lastTransactionRow}`).format.horizontalAlignment = "center";
    sheet.getRange(`E21:E${lastTransactionRow}`).format.font = { name: CJK, size: 11, color: BLACK };
  }
  if (lastPriceRow >= 4) grid(sheet.getRange(`P3:S${lastPriceRow}`));
  if (lastHoldingRow >= 38) {
    sheet.getRange(`P38:Z${lastHoldingRow}`).format.fill = WHITE;
    grid(sheet.getRange(`P37:Z${lastHoldingRow}`));
  }
  sheet.getRange("B21:B212").format.numberFormat = "yyyy-mm-dd";
  sheet.getRange("F21:F212").format.numberFormat = UNITS;
  for (const column of ["G", "H", "I", "J", "K", "M", "N", "O"]) sheet.getRange(`${column}21:${column}212`).format.numberFormat = MONEY;
  sheet.getRange("Q4:Q33").format.numberFormat = MONEY;
  sheet.getRange("S4:S33").format.numberFormat = "yyyy-mm-dd";
  sheet.getRange("R38:R67").format.numberFormat = UNITS;
  sheet.getRange("S38:Y67").format.numberFormat = MONEY;
  setColumnWidths(sheet, {
    A: 19.83, B: 10.66, C: 6.66, D: 16.83, E: 8.33, F: 6.66, G: 11, H: 11.5, I: 12.5, J: 13,
    K: 10.33, L: 4.83, M: 4.5, N: 10.66, O: 13.33, P: 35.16, Q: 17, R: 10.16, S: 9.66,
    T: 8.83, U: 12.5, V: 13, W: 12, X: 13, Y: 13, Z: 13,
  });
  sheet.getRange("A1:AI1").format.rowHeight = 21.75;
  sheet.getRange("A19:AI19").format.rowHeight = 21.75;
  sheet.getRange("A20:AI20").format.rowHeight = 19.5;
  sheet.getRange("A21:AI212").format.rowHeight = 15;
}

function formatAsset(sheet) {
  sheet.showGridLines = false;
  const assetCanvas = sheet.getRange("A1:L201");
  assetCanvas.format.font = { name: BODY, size: 10, color: BLACK };
  assetCanvas.format.verticalAlignment = "center";
  assetCanvas.format.wrapText = false;
  fixedMergedBand(sheet, "A1:H1", "Asset allocation by the four-quadrant method", { font: { name: CJK, size: 12, bold: true, color: WHITE }, horizontalAlignment: "left" });
  lightBand(sheet.getRange("A3:C3"), CJK, 11);
  sheet.getRange("A4:A9").format.font = { name: VALUE, size: 11, bold: true, color: BLACK };
  sheet.getRange("B4:C9").format.font = { name: BODY, size: 10, color: BLACK };
  fixedMergedBand(sheet, "A11:H11", "Asset details", { font: { name: CJK, size: 12, bold: true, color: WHITE }, horizontalAlignment: "left" });
  lightBand(sheet.getRange("A12:H12"), CJK, 11);
  bottomBorder(sheet.getRange("A12:H12"));
  sheet.getRange("J3:K3").format.fill = PURPLE;
  sheet.getRange("J3:K3").format.font = { name: CJK, size: 12, bold: true, color: BLACK };
  sheet.getRange("J3:K3").format.horizontalAlignment = "left";
  sheet.getRange("J4:J7").format.fill = LIGHT_BLUE;
  sheet.getRange("L3:L7").format.fill = LIGHT_BLUE;
  sheet.getRange("K4:K7").format.fill = WHITE;
  for (const address of ["J4:J7", "L3:L7"]) sheet.getRange(address).format.font = { name: CJK, size: 11, bold: true, color: DARK_TEXT };
  sheet.getRange("K4:K7").format.font = { name: BODY, size: 11, color: BLACK };
  sheet.getRange("A13:H201").format.fill = PALE_BLUE;
  sheet.getRange("A13:H201").format.verticalAlignment = "center";
  sheet.getRange("E13:F201").format.fill = "#F3F4F6";
  sheet.getRange("E13:F201").format.font = { name: BODY, size: 10, color: BLACK };
  sheet.getRange("G13:G201").format.font = { name: BODY, size: 10, color: BLACK };
  applyAssetDetailFonts(sheet);
  sheet.getRange("A12:H201").format.horizontalAlignment = "center";
  grid(sheet.getRange("A12:H201"));
  grid(sheet.getRange("J3:L7"));
  sheet.getRange("B4:B9").format.numberFormat = MONEY;
  sheet.getRange("C4:C9").format.numberFormat = "0.0%";
  sheet.getRange("D13:F201").format.numberFormat = MONEY;
  sheet.getRange("H13:H201").format.numberFormat = "yyyy-mm-dd";
  sheet.getRange("A:H").format.columnWidth = 16;
  sheet.getRange("A:A").format.columnWidth = 22;
  sheet.getRange("B:B").format.columnWidth = 18.33;
  sheet.getRange("C:C").format.columnWidth = 10;
  sheet.getRange("D:D").format.columnWidth = 14;
  sheet.getRange("H:H").format.columnWidth = 28;
  sheet.getRange("J:J").format.columnWidth = 14;
  sheet.getRange("K:K").format.columnWidth = 15;
  sheet.getRange("L:L").format.columnWidth = 22;
  sheet.getRange("A1:L1").format.rowHeight = 21.75;
  sheet.getRange("A3:L3").format.rowHeight = 21.75;
  sheet.getRange("A11:L11").format.rowHeight = 21.75;
  sheet.getRange("A12:L12").format.rowHeight = 19.5;
  sheet.getRange("A4:L10").format.rowHeight = 15;
  sheet.getRange("A13:L201").format.rowHeight = 15;
  sheet.getRange("A9:C9").format.fill = "#DDEBF7";
  sheet.getRange("A9:C9").format.font = { name: VALUE, size: 11, bold: true, color: BLACK };
  sheet.getRange("A9:C9").format.borders = { top: { style: "medium", color: GRID } };
}

function formatSummary(sheet) {
  sheet.showGridLines = false;
  const summaryCanvas = sheet.getRange("A1:AZ1046");
  summaryCanvas.format.font = { name: BODY, size: 11, color: BLACK };
  summaryCanvas.format.verticalAlignment = "center";
  summaryCanvas.format.wrapText = false;
  unmergeIfPresent(sheet, "A1:Z1");
  fixedMergedBand(sheet, "A1:AB1", "Net Worth Trend", { font: { name: BODY, size: 14, bold: true, color: WHITE } });
  fixedMergedBand(sheet, "A3:C3", "Trend Summary", { font: { name: BODY, size: 11, bold: true, color: WHITE } });
  unmergeIfPresent(sheet, "M3:N3");
  sheet.getRange("M3:M7").clear({ applyTo: "contents" });
  sheet.getRange("M3:M7").format.fill = WHITE;
  fixedMergedBand(sheet, "N3:O3", "Reconciliation Check", { font: { name: BODY, size: 11, bold: true, color: WHITE } });
  sheet.getRange("A4:A6").format = { fill: LIGHT_BLUE, font: { name: BODY, size: 11, bold: true, color: BLACK }, verticalAlignment: "center" };
  sheet.getRange("N4:N7").format = { fill: LIGHT_BLUE, font: { name: BODY, size: 11, bold: true, color: BLACK }, verticalAlignment: "center" };
  grid(sheet.getRange("A4:B6"));
  grid(sheet.getRange("N4:O7"));
  sheet.getRange("A9").format = { fill: BLUE, font: { name: BODY, size: 11, bold: true, color: WHITE } };
  const linkedTickers = sheet.getRange("A10:AZ10").values[0];
  const lastLinkedTicker = Math.max(0, linkedTickers.reduce((last, value, index) => clean(value) ? index : last, 0));
  grid(sheet.getRange(`A10:${columnLetter(lastLinkedTicker)}11`));
  sheet.getRange(`A10:${columnLetter(lastLinkedTicker)}10`).format.font = { name: BODY, size: 11, bold: true, color: BLACK };
  sheet.getRange(`B10:${columnLetter(lastLinkedTicker)}11`).format.horizontalAlignment = "center";
  sheet.getRange("B4").format.numberFormat = MONEY;
  sheet.getRange("B5").format.numberFormat = "yyyy-mm-dd";
  sheet.getRange("B6").format.numberFormat = MONEY;
  sheet.getRange("O4:O6").format.numberFormat = MONEY;
  const openingLabelIndex = sheet.getRange("A7:AZ7").values[0].findIndex((value) => clean(value) === "Opening Date");
  if (openingLabelIndex >= 0) {
    const openingLabelColumn = columnLetter(openingLabelIndex);
    const openingLabels = sheet.getRange(`${openingLabelColumn}7:${openingLabelColumn}8`);
    openingLabels.format.fill = BLUE;
    openingLabels.format.font = { name: BODY, size: 11, bold: true, color: WHITE };
    openingLabels.format.horizontalAlignment = "center";
    openingLabels.format.verticalAlignment = "center";
  }
  const headers = sheet.getRange("A13:AZ13").values[0];
  const lastHeader = Math.max(0, headers.reduce((last, value, index) => clean(value) ? index : last, 0));
  band(sheet.getRange(`A13:${columnLetter(lastHeader)}13`));
  const lastDateRow = lastPopulatedRow(sheet.getRange("A14:A1046").values, 14);
  if (lastDateRow >= 14) {
    sheet.getRange(`A14:A${lastDateRow}`).format.numberFormat = "yyyy-mm-dd";
    sheet.getRange(`B14:${columnLetter(lastHeader)}${lastDateRow}`).format.horizontalAlignment = "center";
    grid(sheet.getRange(`A13:${columnLetter(lastHeader)}${lastDateRow}`));
  }
  sheet.getRange("A:A").format.columnWidth = 22.83;
  if (lastHeader >= 1) sheet.getRange(`B:${columnLetter(lastHeader)}`).format.columnWidth = 11;
  headers.forEach((value, index) => {
    if (/Stock MV|Total Net Worth/i.test(clean(value))) sheet.getRange(`${columnLetter(index)}:${columnLetter(index)}`).format.columnWidth = 18;
  });
  sheet.getRange("N:N").format.columnWidth = 18;
  sheet.getRange("O:O").format.columnWidth = 18.33;
  sheet.getRange("A1:AZ1").format.rowHeight = 19;
}

function band(range, fontName = BODY, fontSize = 11) {
  colorBand(range, BLUE, WHITE, fontName, fontSize);
}

function lightBand(range, fontName = BODY, fontSize = 11) {
  colorBand(range, LIGHT_BLUE, DARK_TEXT, fontName, fontSize);
}

function colorBand(range, fill, color, fontName = BODY, fontSize = 11) {
  range.format.fill = fill;
  range.format.font = { name: fontName, size: fontSize, bold: true, color };
  range.format.horizontalAlignment = "center";
  range.format.verticalAlignment = "center";
}

function grid(range) {
  range.format.borders = { preset: "all", style: "thin", color: GRID };
}

function setColumnWidths(sheet, widths) {
  for (const [column, width] of Object.entries(widths)) sheet.getRange(`${column}:${column}`).format.columnWidth = width;
}

function applyAssetDetailFonts(sheet) {
  const categories = sheet.getRange("B13:B201").values;
  for (let index = 0; index < categories.length; index += 1) {
    const row = index + 13;
    const category = clean(categories[index][0]);
    if (!category) continue;
    const imported = category === "Account" || category === "Cash";
    const sourceColor = imported ? INPUT_BLUE : LINK_GREEN;
    sheet.getRange(`A${row}`).format.font = { name: CJK, size: 12, color: sourceColor };
    sheet.getRange(`B${row}:C${row}`).format.font = { name: CJK, size: 10, color: sourceColor };
    sheet.getRange(`D${row}`).format.font = { name: BODY, size: 10, color: sourceColor };
    sheet.getRange(`H${row}`).format.font = { name: CJK, size: 10, color: sourceColor };
  }
}

function bottomBorder(range) {
  range.format.borders = { bottom: { style: "thin", color: "#D9D9D9" } };
}

function unmergeIfPresent(sheet, address) {
  try { sheet.unmergeCells(address); } catch {}
}

function fixedMergedBand(sheet, address, title, options = {}) {
  unmergeIfPresent(sheet, address);
  sheet.mergeCells(address);
  const range = sheet.getRange(address);
  range.values = [[title]];
  band(range, options.font?.name ?? CJK, options.font?.size ?? 12);
  range.format.font = options.font ?? { name: CJK, size: 12, bold: true, color: WHITE };
  range.format.horizontalAlignment = options.horizontalAlignment ?? "center";
  range.format.verticalAlignment = "center";
}

function fixedMergedTitle(sheet, address, title, font, fill = WHITE) {
  unmergeIfPresent(sheet, address);
  sheet.mergeCells(address);
  const range = sheet.getRange(address);
  range.values = [[title]];
  range.format.fill = fill;
  range.format.font = font;
  range.format.horizontalAlignment = "left";
  range.format.verticalAlignment = "center";
}

function lastPopulatedRow(values, firstRow) {
  let result = firstRow - 1;
  values.forEach((row, index) => { if (clean(row[0])) result = firstRow + index; });
  return result;
}

function clean(value) {
  return value === null || value === undefined ? "" : String(value).trim();
}

function columnLetter(index) {
  let value = index + 1;
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

async function exists(file) {
  try { await fs.access(file); return true; } catch { return false; }
}

function fail(message) {
  throw new Error(message);
}

async function main() {
  const args = parseCliArgs(process.argv.slice(2));
  if (!args.workbook || !args.output) fail("Usage: apply_workbook_format.mjs --workbook FILE --output FILE [--allow-existing-output]");
  console.log(JSON.stringify(await runWorkbookFormat({
    workbookPath: args.workbook,
    outputPath: args.output,
    allowExistingOutput: Boolean(args.allowExistingOutput),
  }), null, 2));
}

function parseCliArgs(argv) {
  const args = {};
  for (let index = 0; index < argv.length; index += 1) {
    const key = argv[index];
    if (key === "--allow-existing-output") args.allowExistingOutput = true;
    else if (key.startsWith("--")) args[key.slice(2).replace(/-([a-z])/g, (_, character) => character.toUpperCase())] = argv[++index];
    else fail(`Unexpected argument: ${key}`);
  }
  return args;
}

const directRun = process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href;
if (directRun) main().catch((error) => { console.error(error instanceof Error ? error.message : String(error)); process.exitCode = 1; });

#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import { createRequire } from "node:module";
import { pathToFileURL } from "node:url";
import { saveWorkbookFormulaSafe } from "./workbook_io.mjs";
import { applyWorkbookFormat } from "./apply_workbook_format.mjs";

const dependencies = process.env.CODEX_NODE_MODULES;
if (!dependencies) throw new Error("CODEX_NODE_MODULES is required");
const require = createRequire(path.join(dependencies, "package.json"));
const { FileBlob, SpreadsheetFile, Workbook } = require("@oai/artifact-tool");

const BLUE = "#1F4E78";
const LIGHT_BLUE = "#D9EAF7";
const MONEY = "#,##0.00;[Red](#,##0.00)";
const ACCOUNTS = [
  ["WeChat", "Account", "CNY", "Security"], ["中信银行", "Account", "USD", "Security"],
  ["中信银行", "Account", "CNY", "Security"], ["支付宝", "Account", "CNY", "Security"],
  ["Chase", "Account", "USD", "Security"], ["Cash", "Cash", "USD", "Security"],
  ["Robinhood Cash", "Cash", "USD", "Security"], ["Fidelity Cash", "Cash", "USD", "Security"],
  ["IBKR Cash", "Cash", "USD", "Security"],
  ["Charles Cash", "Cash", "USD", "Security"],
];

export async function createTreasuryWorkbook({ inputFile, outputPath }) {
  const config = await readBootstrap(inputFile);
  const workbook = Workbook.create();
  buildCashflow(workbook, config);
  buildStock(workbook, config);
  buildAsset(workbook, config);
  buildSummary(workbook, config);
  applyWorkbookFormat(workbook);
  await fs.mkdir(path.dirname(path.resolve(outputPath)), { recursive: true });
  await saveWorkbookFormulaSafe(workbook, path.resolve(outputPath));
  return { output: path.resolve(outputPath), sheets: ["CashFlow", "Stock", "Asset", "Summary"], ...config.report };
}

export async function continueTreasuryWorkbook({ bootstrapWorkbookPath, outputPath }) {
  const source = path.resolve(bootstrapWorkbookPath);
  const destination = path.resolve(outputPath);
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(source));
  const requiredSheets = ["CashFlow", "Stock", "Asset", "Summary"];
  const presentSheets = listSheetNames(workbook);
  const missingSheets = requiredSheets.filter((name) => !presentSheets.includes(name));
  if (missingSheets.length) throw new Error(`Bootstrap workbook is missing required sheets: ${missingSheets.join(", ")}`);
  for (const name of presentSheets.reverse()) {
    if (!requiredSheets.includes(name)) workbook.worksheets.getItem(name).delete();
  }
  migrateLegacyStockLayout(workbook.worksheets.getItem("Stock"));
  migrateLegacyAssetAccounts(workbook.worksheets.getItem("Asset"));
  await fs.mkdir(path.dirname(destination), { recursive: true });
  await saveWorkbookFormulaSafe(workbook, destination);
  return { output: destination, sheets: requiredSheets, mode: "continue", bootstrapWorkbook: source };
}

// Older verified bootstrap files stored the holdings summary immediately below
// the price input block (P20:Z34).  The current four-sheet layout reserves the
// same block at P36:Z67 so that new securities can be added without colliding
// with the transaction ledger.  Normalize that legacy block during an explicit
// continue before any importer reads the fixed current-layout addresses.
function migrateLegacyStockLayout(sheet) {
  const currentHeader = sheet.getRange("P37:Z37").values[0].map((value) => String(value ?? "").trim());
  const legacyHeader = sheet.getRange("P21:Z21").values[0].map((value) => String(value ?? "").trim());
  const expected = ["Ticker", "名称", "Total Units", "Total Cost", "Avg Cost", "Current Price", "Current Value", "未实现盈亏", "已实现盈亏", "总收益", "Bucket"];
  if (JSON.stringify(currentHeader) === JSON.stringify(expected)) return;
  if (JSON.stringify(legacyHeader) !== JSON.stringify(expected)) return;

  // Copy only the legacy section (rows 20:35) so the source and destination do
  // not overlap; copyFrom preserves formats and adjusts relative formulas.
  sheet.getRange("P36:Z51").copyFrom(sheet.getRange("P20:Z35"), "all");
  sheet.getRange("P20:Z35").clear({ applyTo: "all" });
  sheet.getRange("B3").formulas = [["=IF($B$9=\"USD\",SUM($S$38:$S$67),SUM($S$38:$S$67)*$B$10)"]];
  sheet.getRange("B4").formulas = [["=IF($B$9=\"USD\",SUM($V$38:$V$67),SUM($V$38:$V$67)*$B$10)"]];
  sheet.getRange("B5").formulas = [["=B4-B3"]];
  sheet.getRange("B6").formulas = [["=COUNTIF($P$38:$P$67,\"<>\")"]];
}

function migrateLegacyAssetAccounts(sheet) {
  const headers = sheet.getRange("A12:H12").values[0].map((value) => String(value ?? "").trim());
  const expected = ["资产名称", "Category", "币种", "Amount", "FX Rate", "Converted Amount", "Bucket", "Update time"];
  if (JSON.stringify(headers) !== JSON.stringify(expected)) return;
  const rows = sheet.getRange("A13:C201").values;
  const names = new Set(rows.map((row) => String(row[0] ?? "").trim()));
  const templateRow = 21;
  let targetRow = 13 + rows.findIndex((row) => !String(row[0] ?? "").trim());
  if (targetRow < 13) targetRow = 202;
  for (const [name, category, currency, bucket] of [["IBKR Cash", "Cash", "USD", "Security"], ["Charles Cash", "Cash", "USD", "Security"]]) {
    if (names.has(name)) continue;
    if (targetRow > 201) throw new Error(`Asset detail block has no row for legacy account ${name}.`);
    sheet.getRange(`A${targetRow}:H${targetRow}`).copyFrom(sheet.getRange(`A${templateRow}:H${templateRow}`), "all");
    sheet.getRange(`A${targetRow}:C${targetRow}`).values = [[name, category, currency]];
    sheet.getRange(`G${targetRow}`).values = [[bucket]];
    sheet.getRange(`D${targetRow}`).values = [[null]];
    sheet.getRange(`H${targetRow}`).values = [[null]];
    names.add(name);
    targetRow += 1;
  }
  for (let row = 4; row <= 8; row += 1) {
    sheet.getRange(`B${row}`).formulas = [[`=SUMIFS($F$13:$F$201,$G$13:$G$201,A${row})`]];
    sheet.getRange(`C${row}`).formulas = [[`=IFERROR(B${row}/$B$9,0)`]];
  }
  sheet.getRange("B9").formulas = [["=SUM(B4:B8)"]];
  sheet.getRange("C9").formulas = [["=IF(B9=0,0,SUM(C4:C8))"]];
}

function buildCashflow(workbook) {
  const sheet = workbook.worksheets.add("CashFlow");
  sheet.showGridLines = false;
  sheet.getRange("A1:AM500").format.font = { name: "Calibri", size: 11 };
  sheet.getRange("A1:AM500").format.horizontalAlignment = "center";
  sheet.getRange("A1:AM500").format.verticalAlignment = "center";
  section(sheet, "A1:L1", "Summary");
  sheet.getRange("A2:B2").values = [["Selected Month:", null]];
  sheet.getRange("A3:K5").values = [
    ["Total Income (Currency)", null, null, "Start Date", null, null, "End Date", null, null, "Income Record Count", null],
    ["Total Expense (Currency)", null, null, "Start Date", null, null, "End Date", null, null, "Expense Record Count", null],
    ["Total Transfer (Currency)", null, null, "Start Date", null, null, "End Date", null, null, "Transfer Record Count", null],
  ];
  sheet.getRange("A7:B7").values = [["Net Cash Flow", null]];
  section(sheet, "A10:J10", "Income"); section(sheet, "N10:W10", "Expenses"); section(sheet, "AA10:AM10", "Transfer");
  const income = ["Date", "Year-Month", "Category", "Tag", "Description", "Amount", "Currency", "Account", "FX Rate", "Converted Amount"];
  const expense = [...income];
  const transfer = ["Date", "Year-Month", "Counterparty", "Category", "Tag", "Description", "Amount", "Currency", "Account", "FX Rate", "Converted Amount", "Outstanding", "Status"];
  header(sheet.getRange("A11:J11"), income); header(sheet.getRange("N11:W11"), expense); header(sheet.getRange("AA11:AM11"), transfer);
  sheet.getRange("B3").formulas = [["=SUMIFS($J$12:$J$454,$B$12:$B$454,$B$2)"]];
  sheet.getRange("B4").formulas = [["=SUMIFS($W$12:$W$456,$O$12:$O$456,$B$2)"]];
  sheet.getRange("B5").formulas = [["=SUMIFS($AK$12:$AK$500,$AB$12:$AB$500,$B$2)"]];
  sheet.getRange("E3").formulas = [["=IF(K3=0,\"\",_xlfn.MINIFS($A$12:$A$454,$B$12:$B$454,$B$2))"]];
  sheet.getRange("H3").formulas = [["=IF(K3=0,\"\",_xlfn.MAXIFS($A$12:$A$454,$B$12:$B$454,$B$2))"]];
  sheet.getRange("K3").formulas = [["=COUNTIF($B$12:$B$454,$B$2)"]];
  sheet.getRange("E4").formulas = [["=IF(K4=0,\"\",_xlfn.MINIFS($N$12:$N$456,$O$12:$O$456,$B$2))"]];
  sheet.getRange("H4").formulas = [["=IF(K4=0,\"\",_xlfn.MAXIFS($N$12:$N$456,$O$12:$O$456,$B$2))"]];
  sheet.getRange("K4").formulas = [["=COUNTIF($O$12:$O$456,$B$2)"]];
  sheet.getRange("E5").formulas = [["=IF(K5=0,\"\",_xlfn.MINIFS($AA$12:$AA$500,$AB$12:$AB$500,$B$2))"]];
  sheet.getRange("H5").formulas = [["=IF(K5=0,\"\",_xlfn.MAXIFS($AA$12:$AA$500,$AB$12:$AB$500,$B$2))"]];
  sheet.getRange("K5").formulas = [["=COUNTIF($AB$12:$AB$500,$B$2)"]];
  sheet.getRange("B7").formulas = [["=B3-B4"]];
  for (const range of ["A12:A454", "N12:N456", "AA12:AA500"]) sheet.getRange(range).format.numberFormat = "yyyy-mm-dd";
  for (const range of ["F12:F454", "J12:J454", "S12:S456", "W12:W456", "AG12:AG500", "AK12:AL500"]) sheet.getRange(range).format.numberFormat = MONEY;
  for (const col of ["A:J", "N:W", "AA:AM"]) sheet.getRange(col).format.columnWidth = 13;
  sheet.getRange("E:E").format.columnWidth = 20; sheet.getRange("R:R").format.columnWidth = 20; sheet.getRange("AF:AF").format.columnWidth = 20;
}

function buildStock(workbook, config) {
  const sheet = workbook.worksheets.add("Stock");
  sheet.showGridLines = false;
  sheet.getRange("A1:AI263").format.font = { name: "Calibri", size: 11 };
  sheet.getRange("A1").values = [["Stock buying, selling, and holding analysis"]];
  sheet.getRange("A3:B6").values = [["总投入成本（主货币）", null], ["当前总市值（主货币）", null], ["浮动盈亏（主货币）", null], ["持仓品种数", null]];
  sheet.getRange("B3").formulas = [["=IF($B$9=\"USD\",SUM($S$38:$S$67),SUM($S$38:$S$67)*$B$10)"]];
  sheet.getRange("B4").formulas = [["=IF($B$9=\"USD\",SUM($V$38:$V$67),SUM($V$38:$V$67)*$B$10)"]];
  sheet.getRange("B5").formulas = [["=B4-B3"]];
  sheet.getRange("B6").formulas = [["=COUNTIF($P$38:$P$67,\"<>\")"]];
  sheet.getRange("A9:B10").values = [["主货币", config.baseCurrency], ["USD/CNY", config.usdCny]];
  section(sheet, "P2:S2", "当前价格输入（可手动更新）");
  header(sheet.getRange("P3:S3"), ["Ticker", "Current Price", "币种", "更新日期"]);
  section(sheet, "A19:O19", "交易明细");
  header(sheet.getRange("A20:O20"), ["Broker", "Date", "Ticker", "Name", "Action", "Units", "Trade Price", "Trade Value", "Current Price", "Current Value", "未实现盈亏", "币种", "Fee", "Cost Basis", "已实现盈亏"]);
  section(sheet, "P36:Z36", "持仓汇总（按当前市值）");
  header(sheet.getRange("P37:Z37"), ["Ticker", "名称", "Total Units", "Total Cost", "Avg Cost", "Current Price", "Current Value", "未实现盈亏", "已实现盈亏", "总收益", "Bucket"]);
  sheet.getRange("A21:O212").format.fill = "#F8FBFF"; sheet.getRange("P38:Z67").format.fill = "#F8FBFF";
  sheet.getRange("A21:O212").format.horizontalAlignment = "center"; sheet.getRange("P3:Z67").format.horizontalAlignment = "center";
  sheet.getRange("B21:B212").format.numberFormat = "yyyy-mm-dd"; sheet.getRange("S4:S33").format.numberFormat = "yyyy-mm-dd";
  sheet.getRange("F21:F212").format.numberFormat = "0.000000";
  for (const range of ["G21:K212", "M21:O212", "Q4:Q33", "S38:Y67"]) sheet.getRange(range).format.numberFormat = MONEY;
  sheet.getRange("A:Z").format.columnWidth = 13; sheet.getRange("D:D").format.columnWidth = 24; sheet.getRange("Q:Q").format.columnWidth = 24;
}

function buildAsset(workbook, config) {
  const sheet = workbook.worksheets.add("Asset");
  sheet.showGridLines = false;
  sheet.getRange("A1:L201").format.font = { name: "Calibri", size: 11 };
  section(sheet, "A1:H1", "Asset allocation by the four-quadrant method");
  header(sheet.getRange("A3:C3"), ["Category", "Amount (Currency)", "Weight"]);
  sheet.getRange("A4:A9").values = [["Security"], ["Growth"], ["High Growth"], ["Luxury"], ["Unclassified"], ["Total"]];
  for (let row = 4; row <= 8; row += 1) {
    const label = sheet.getRange(`A${row}`).values[0][0];
    sheet.getRange(`B${row}`).formulas = [[`=SUMIFS($F$13:$F$201,$G$13:$G$201,A${row})`]];
    sheet.getRange(`C${row}`).formulas = [[`=IFERROR(B${row}/$B$9,0)`]];
  }
  sheet.getRange("B9").formulas = [["=SUM(B4:B8)"]]; sheet.getRange("C9").formulas = [["=IF(B9=0,0,SUM(C4:C8))"]];
  sheet.getRange("J3:K4").values = [["主货币", config.baseCurrency], ["USD/CNY", config.usdCny]];
  sheet.getRange("L3:L7").values = [["在 USD / CNY 间切换"], ["1 USD = ? CNY"], ["CNY → 主货币"], ["USD → 主货币"], ["汇总按主货币显示"]];
  sheet.getRange("J5:J7").values = [["CNY"], ["USD"], ["状态"]];
  sheet.getRange("K5").formulas = [["=IF($K$3=\"CNY\",1,1/$K$4)"]];
  sheet.getRange("K6").formulas = [["=IF($K$3=\"USD\",1,$K$4)"]];
  sheet.getRange("K7").formulas = [["=IF(OR($K$3=\"\",$K$4=\"\"),\"⚠ 请设置主货币与汇率\",\"OK\")"]];
  section(sheet, "A11:H11", "Asset details");
  header(sheet.getRange("A12:H12"), ["资产名称", "Category", "币种", "Amount", "FX Rate", "Converted Amount", "Bucket", "Update time"]);
  sheet.getRange("A13:H201").format.fill = "#EAF3FF";
  ACCOUNTS.forEach((item, index) => {
    const row = 13 + index;
    sheet.getRange(`A${row}:C${row}`).values = [[item[0], item[1], item[2]]]; sheet.getRange(`G${row}`).values = [[item[3]]];
    sheet.getRange(`E${row}`).formulas = [[`=IF(C${row}=$K$3,1,IF(AND(C${row}="USD",$K$3="CNY"),$K$4,IF(AND(C${row}="CNY",$K$3="USD"),1/$K$4,"")))`]];
    sheet.getRange(`F${row}`).formulas = [[`=IF(OR(D${row}="",E${row}=""),"",D${row}*E${row})`]];
  });
  sheet.getRange("D13:F201").format.numberFormat = MONEY; sheet.getRange("H13:H201").format.numberFormat = "yyyy-mm-dd";
  sheet.getRange("A:H").format.columnWidth = 16; sheet.getRange("A:A").format.columnWidth = 20;
  const allocationChart = sheet.charts.add("pie", sheet.getRange("A3:B7"), "Auto");
  allocationChart.title.text = "当前资产分配";
  allocationChart.hasLegend = true;
  allocationChart.setPosition(sheet.getRange("J11:P25"));
  allocationChart.width = 520;
  allocationChart.height = 300;
}

function buildSummary(workbook, config) {
  const sheet = workbook.worksheets.add("Summary");
  sheet.showGridLines = false;
  sheet.getRange("A1:AZ1046").format.font = { name: "Calibri", size: 11 };
  section(sheet, "A1:Z1", "Net Worth Trend");
  section(sheet, "A3:C3", "Trend Summary");
  sheet.getRange("A4:A6").values = [["Latest Total Net Worth"], ["Latest Date"], ["Latest Stock MV"]];
  section(sheet, "N3:O3", "Reconciliation Check");
  sheet.getRange("N4:N7").values = [["Real Asset"], ["Calculated Asset"], ["Total Difference"], ["Status"]];
  sheet.getRange("O4").formulas = [["='Asset'!B9"]];
  sheet.getRange("O5").formulas = [["=B4"]];
  sheet.getRange("O6").formulas = [["=O4-O5"]];
  sheet.getRange("O7").formulas = [["=IF(ABS(O6)<100,\"OK\",\"Check\")"]];
  sheet.getRange("A9").values = [["Linked Holdings"]];
  sheet.getRange("A10:D13").values = [["Ticker", null, null, null], ["Units Held", null, null, null], [null, null, null, null], ["Date", "Stock MV (USD)", "Stock MV (Base)", "Total Net Worth (Base)"]];
  header(sheet.getRange("A13:D13"), ["Date", "Stock MV (USD)", "Stock MV (Base)", "Total Net Worth (Base)"]);
  sheet.getRange("F12:F13").values = [["实时持有数量"], ["实时持有现金"]];
  sheet.getRange("F7:G8").values = [["Opening Date", config.openingDate], ["Opening Cash Balance (Base)", config.openingCash]];
  sheet.getRange("G7").format.numberFormat = "yyyy-mm-dd"; sheet.getRange("G8").format.numberFormat = MONEY;
  sheet.getRange("A14:D1046").format.fill = "#EAF3FF"; sheet.getRange("A14:A1046").format.numberFormat = "yyyy-mm-dd"; sheet.getRange("B14:D1046").format.numberFormat = MONEY;
  sheet.getRange("A:A").format.columnWidth = 12; sheet.getRange("B:D").format.columnWidth = 21; sheet.getRange("F:G").format.columnWidth = 18;
  sheet.charts.add("line", { title: "Net Worth Over Time", categories: [""], series: [{ name: "Total Net Worth", values: [0] }], hasLegend: false, from: { row: 1, col: 9 }, extent: { widthPx: 700, heightPx: 340 } });
}

function section(sheet, address, title) { const range = sheet.getRange(address); range.merge(); range.values = [[title]]; range.format.fill = BLUE; range.format.font = { name: "Calibri", size: 11, bold: true, color: "#FFFFFF" }; range.format.horizontalAlignment = "center"; }
function header(range, values) { range.values = [values]; range.format.fill = BLUE; range.format.font = { name: "Calibri", size: 11, bold: true, color: "#FFFFFF" }; range.format.horizontalAlignment = "center"; range.format.verticalAlignment = "center"; }

function listSheetNames(workbook) {
  const names = [];
  for (let index = 0; ; index += 1) {
    try {
      const sheet = workbook.worksheets.getItemAt(index);
      if (!sheet) break;
      names.push(sheet.name);
    } catch {
      break;
    }
  }
  return names;
}

export async function readBootstrap(file) {
  const text = await fs.readFile(file, "utf8");
  const value = (label) => text.match(new RegExp(`^${label}:\\s*(.+?)\\s*$`, "mi"))?.[1]?.trim();
  const openingDateText = value("Opening Date"); const openingCash = Number(value("Opening Cash Balance \\(Base\\)"));
  const baseCurrency = value("Base Currency")?.toUpperCase(); const usdCny = Number(value("USD/CNY"));
  if (!/^\d{4}-\d{2}-\d{2}$/.test(openingDateText || "")) throw new Error("monthly-close-input.md needs Opening Date as YYYY-MM-DD.");
  if (!Number.isFinite(openingCash) || !["USD", "CNY"].includes(baseCurrency) || !Number.isFinite(usdCny) || usdCny <= 0) throw new Error("Invalid workbook bootstrap values in monthly-close-input.md.");
  const lines = text.split(/\r?\n/);
  const headerIndex = lines.findIndex((line) => /^\s*\|\s*Ticker\s*\|\s*Units\s*\|\s*Total Cost\s*\|\s*Realized P\/L\s*\|\s*Bucket\s*\|\s*$/i.test(line));
  if (headerIndex < 0) throw new Error("monthly-close-input.md needs the opening holdings table.");
  const openingHoldings = [];
  const seen = new Set();
  for (let index = headerIndex + 2; index < lines.length; index += 1) {
    const line = lines[index].trim();
    if (!line.startsWith("|")) break;
    const cells = line.slice(1, line.endsWith("|") ? -1 : undefined).split("|").map((cell) => cell.trim());
    if (cells.length < 5) continue;
    const ticker = cells[0].toUpperCase() === "ECHO" ? "SATS" : cells[0].toUpperCase();
    const units = Number(cells[1]);
    const totalCost = Number(cells[2]);
    const realized = Number(cells[3]);
    const bucket = cells[4];
    if (!/^[A-Z0-9.-]+$/.test(ticker) || !Number.isFinite(units) || !Number.isFinite(totalCost) || !Number.isFinite(realized) || !bucket) throw new Error(`Invalid opening holding at row ${index + 1}.`);
    if (seen.has(ticker)) throw new Error(`Duplicate opening holding ticker: ${ticker}`);
    seen.add(ticker);
    openingHoldings.push({ ticker, units, totalCost, realized, bucket });
  }
  if (!openingHoldings.length) throw new Error("monthly-close-input.md opening holdings table is empty.");
  return {
    openingDate: new Date(`${openingDateText}T00:00:00.000Z`),
    openingDateText,
    openingCash,
    openingHoldings,
    baseCurrency,
    usdCny,
    report: { openingDate: openingDateText, openingCash, openingHoldings, baseCurrency, usdCny },
  };
}

async function main() { const args = Object.fromEntries(process.argv.slice(2).reduce((a, v, i, x) => v.startsWith("--") ? [...a, [v.slice(2), x[i + 1]]] : a, [])); if (!args.output) throw new Error("Usage: create_workbook.mjs --output FILE [--input-file monthly-close-input.md]"); console.log(JSON.stringify(await createTreasuryWorkbook({ inputFile: path.resolve(args["input-file"] || "monthly-close-input.md"), outputPath: args.output }), null, 2)); }
const directRun = process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href;
if (directRun) main().catch((error) => { console.error(error.message); process.exitCode = 1; });

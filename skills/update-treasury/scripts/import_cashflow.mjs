#!/usr/bin/env node

import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { createRequire } from "node:module";
import { pathToFileURL } from "node:url";
import { createTreasuryWorkbook } from "./create_workbook.mjs";
import { saveWorkbookFormulaSafe } from "./workbook_io.mjs";

const dependencies = process.env.CODEX_NODE_MODULES;
if (!dependencies) throw new Error("CODEX_NODE_MODULES is required");
const require = createRequire(path.join(dependencies, "package.json"));
const { FileBlob, SpreadsheetFile, Workbook } = require("@oai/artifact-tool");
const JSZip = require("jszip");

const REQUIRED_CASHFLOW_HEADERS = ["日期", "分类", "备注", "类型", "金额", "货币", "账户"];
const SUPPORTED_ACCOUNTS = new Set([
  "Chase",
  "WeChat",
  "中信银行",
  "支付宝",
  "Cash",
  "Robinhood Cash",
  "Fidelity Cash",
  "IBKR Cash",
  "Charles Cash",
]);
const ACCOUNT_CURRENCIES = new Map([
  ["Chase", "USD"],
  ["WeChat", "CNY"],
  ["支付宝", "CNY"],
  ["Cash", "USD"],
  ["Robinhood Cash", "USD"],
  ["Fidelity Cash", "USD"],
  ["IBKR Cash", "USD"],
  ["Charles Cash", "USD"],
]);

const CASHFLOW_BLOCKS = {
  income: { startRow: 12, endRow: 454, startCol: 0, sourceWidth: 8 },
  expense: { startRow: 12, endRow: 456, startCol: 13, sourceWidth: 8 },
  transfer: { startRow: 12, endRow: 500, startCol: 26, sourceWidth: 9 },
};

const COMPACT_HEADERS = {
  income: ["Date", "Year-Month", "Category", "Tag", "Description", "Amount", "Currency", "Account", "FX Rate", "Converted Amount"],
  expense: ["Date", "Year-Month", "Category", "Tag", "Description", "Amount", "Currency", "Account", "FX Rate", "Converted Amount"],
  transfer: ["Date", "Year-Month", "Counterparty", "Category", "Tag", "Description", "Amount", "Currency", "Account", "FX Rate", "Converted Amount", "Outstanding", "Status"],
};

export async function discoverCashflowCsvs(cashflowDir) {
  const directory = path.resolve(cashflowDir);
  const entries = await fs.readdir(directory, { withFileTypes: true });
  const recognized = [];
  const ignored = [];

  for (const entry of entries.sort((a, b) => a.name.localeCompare(b.name))) {
    if (!entry.isFile() || !entry.name.toLowerCase().endsWith(".csv") || entry.name.startsWith("~$")) continue;
    const file = path.join(directory, entry.name);
    const { headers } = await readCsvRecords(file);
    if (REQUIRED_CASHFLOW_HEADERS.every((header) => headers.includes(header))) recognized.push(file);
    else ignored.push(file);
  }

  if (!recognized.length) fail(`No recognized CashFlow CSV files found in ${directory}.`);
  return { directory, recognized, ignored };
}

export async function importCashflow(workbook, csvPath, { fxRates = null, throughDate = null } = {}) {
  const rows = await readCsvRecords(csvPath);
  requireHeaders(rows.headers, REQUIRED_CASHFLOW_HEADERS);
  const cutoff = throughDate ? parseDate(throughDate) : null;
  if (throughDate && !cutoff) fail(`Invalid CashFlow cutoff date: ${throughDate}`);
  const income = [];
  const expense = [];
  const transfer = [];
  let transferSourceRecords = 0;
  const rejected = [];
  const excludedAfterCutoff = [];
  const acceptedDates = [];

  for (const raw of rows.records) {
    const type = clean(raw["类型"]);
    const date = parseDate(clean(raw["日期"]));
    if (date && cutoff && date.text > cutoff.text) {
      excludedAfterCutoff.push({ row: raw.__row, date: date.text, type });
      continue;
    }
    const amount = parseNumber(raw["金额"], "金额");
    const currency = clean(raw["货币"]).toUpperCase();
    const account = normalizeAccount(raw["账户"]);

    if (!date || !["USD", "CNY"].includes(currency) || !SUPPORTED_ACCOUNTS.has(account)) {
      rejected.push({ row: raw.__row, reason: "unsupported date, currency, or account" });
      continue;
    }

    acceptedDates.push(date.text);
    if (type === "收入") {
      income.push({
        key: keyOf([date.text, clean(raw["备注"]) || clean(raw["分类"]), clean(raw["详细备注"]) || null, amount, currency, account]),
        values: [date.value, date.month, clean(raw["备注"]) || clean(raw["分类"]), null, clean(raw["详细备注"]) || null, amount, currency, account],
      });
    } else if (type === "支出") {
      expense.push({
        key: keyOf([date.text, clean(raw["分类"]), clean(raw["备注"]) || null, clean(raw["详细备注"]) || null, amount, currency, account]),
        values: [date.value, date.month, clean(raw["分类"]), clean(raw["备注"]) || null, clean(raw["详细备注"]) || null, amount, currency, account],
      });
    } else if (["转账", "转入", "转出"].includes(type)) {
      const counterparty = clean(raw["备注"]);
      if (!counterparty) {
        rejected.push({ row: raw.__row, reason: "transfer requires Counterparty in 备注" });
        continue;
      }
      const signedAmount = type === "转出" ? -Math.abs(amount) : type === "转入" ? Math.abs(amount) : amount;
      transfer.push({
        key: keyOf([date.text, counterparty, clean(raw["分类"]) || null, clean(raw["详细备注"]) || null, signedAmount, currency, account]),
        values: [date.value, date.month, counterparty, clean(raw["分类"]) || null, null, clean(raw["详细备注"]) || null, signedAmount, currency, account],
      });
      transferSourceRecords += 1;
    } else {
      rejected.push({ row: raw.__row, reason: `unsupported transaction type: ${type || "blank"}` });
    }
  }

  for (const raw of rows.transferRecords) {
    const date = parseDate(clean(raw["日期"]));
    if (date && cutoff && date.text > cutoff.text) {
      excludedAfterCutoff.push({ row: raw.__row, date: date.text, type: "transfer" });
      continue;
    }
    const sourceAccount = normalizeAccount(raw["转出账户"]);
    const destinationAccount = normalizeAccount(raw["转入账户"]);
    const amount = Math.abs(parseNumber(raw["金额"], "transfer amount"));
    const sourceExternal = sourceAccount === "外部";
    const destinationExternal = destinationAccount === "外部";
    const sourceCurrency = ACCOUNT_CURRENCIES.get(sourceAccount);
    const destinationCurrency = ACCOUNT_CURRENCIES.get(destinationAccount);
    if (date && sourceExternal !== destinationExternal) {
      const account = sourceExternal ? destinationAccount : sourceAccount;
      const currency = ACCOUNT_CURRENCIES.get(account);
      if (!SUPPORTED_ACCOUNTS.has(account) || !currency) {
        rejected.push({ row: raw.__row, reason: "unsupported external-transfer account" });
        continue;
      }
      acceptedDates.push(date.text);
      const sourceNote = clean(raw["备注"]);
      const signedAmount = sourceExternal ? amount : -amount;
      transfer.push({
        key: keyOf([date.text, "外部", "External Transfer", sourceNote || "External transfer", signedAmount, currency, account, null]),
        values: [date.value, date.month, "外部", "External Transfer", null, sourceNote || "External transfer", signedAmount, currency, account],
      });
      transferSourceRecords += 1;
      continue;
    }
    if (!date || !SUPPORTED_ACCOUNTS.has(sourceAccount) || !SUPPORTED_ACCOUNTS.has(destinationAccount) || !sourceCurrency || sourceCurrency !== destinationCurrency) {
      rejected.push({ row: raw.__row, reason: "unsupported transfer date, account, or cross-currency pair" });
      continue;
    }

    acceptedDates.push(date.text);
    const sourceNote = clean(raw["备注"]);
    const counterparty = sourceNote || `${sourceAccount} → ${destinationAccount}`;
    const description = sourceNote || "Internal transfer";
    for (const [account, signedAmount] of [[sourceAccount, -amount], [destinationAccount, amount]]) {
      transfer.push({
        key: keyOf([date.text, counterparty, "Internal Transfer", description, signedAmount, sourceCurrency, account, null]),
        values: [date.value, date.month, counterparty, "Internal Transfer", null, description, signedAmount, sourceCurrency, account],
      });
    }
    transferSourceRecords += 1;
  }

  if (rejected.length) fail(`CashFlow CSV has rejected rows: ${JSON.stringify(rejected.slice(0, 20))}`);

  const byDateAscending = (a, b) => excelDateText(a.values[0]).localeCompare(excelDateText(b.values[0]));
  income.sort(byDateAscending);
  expense.sort(byDateAscending);
  transfer.sort(byDateAscending);

  const sheet = ensureCompactCashflowLayout(workbook);
  assertHeaders(sheet, "A11:J11", COMPACT_HEADERS.income);
  assertHeaders(sheet, "N11:W11", COMPACT_HEADERS.expense);
  assertHeaders(sheet, "AA11:AM11", COMPACT_HEADERS.transfer);

  return {
    source: path.resolve(csvPath),
    sourceRows: rows.records.length + rows.transferRecords.length,
    excludedAfterCutoff,
    latestDate: acceptedDates.sort().at(-1) || null,
    income: { sourceRecords: income.length, ...appendCashBlock(sheet, income, "income", fxRates) },
    expense: { sourceRecords: expense.length, ...appendCashBlock(sheet, expense, "expense", fxRates) },
    transfer: { sourceRecords: transferSourceRecords, ...appendCashBlock(sheet, transfer, "transfer", fxRates) },
  };
}

export async function runCashflowImport({ workbookPath, outputPath, cashflowDir, cashflowCsvs = [], throughDate = null, dryRun = false, allowExistingOutput = false, cashflowOnly = false }) {
  const sourceWorkbook = path.resolve(workbookPath);
  const destinationWorkbook = path.resolve(outputPath);
  if (sourceWorkbook === destinationWorkbook) fail("Refusing to overwrite the source workbook. Choose a distinct output path.");
  if (!dryRun && !allowExistingOutput && await exists(destinationWorkbook)) fail(`Output already exists: ${destinationWorkbook}`);

  let discovery = null;
  let sources = cashflowCsvs.map((file) => path.resolve(file));
  if (!sources.length) {
    discovery = await discoverCashflowCsvs(cashflowDir);
    sources = discovery.recognized;
  }

  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(sourceWorkbook));
  const cashflowSheet = ensureCompactCashflowLayout(workbook);
  const fxRates = cashflowOnly ? inferCashflowFxRates(cashflowSheet) : null;
  const reports = [];
  for (const source of sources) reports.push(await importCashflow(workbook, source, { fxRates, throughDate }));
  const latestAcceptedDate = reports.map((report) => report.latestDate).filter(Boolean).sort().at(-1) || null;
  if (latestAcceptedDate) cashflowSheet.getRange("B2").values = [[latestAcceptedDate.slice(0, 7)]];
  if (cashflowOnly) {
    refreshStandaloneTransferStatus(workbook.worksheets.getItem("CashFlow"));
    prepareCashflowOnlyPreview(workbook);
  }

  if (!dryRun) {
    await fs.mkdir(path.dirname(destinationWorkbook), { recursive: true });
    await saveCashflowWorkbook(workbook, destinationWorkbook);
  }

  return {
    workbook: sourceWorkbook,
    output: dryRun ? null : destinationWorkbook,
    cashflowDirectory: discovery?.directory || null,
    recognizedFiles: sources,
    ignoredFiles: discovery?.ignored || [],
    cashflowOnly,
    files: reports,
    totals: aggregateReports(reports),
  };
}

export async function saveCashflowWorkbook(workbook, destinationWorkbook) {
  await saveWorkbookFormulaSafe(workbook, destinationWorkbook);
  const exportedBytes = await fs.readFile(destinationWorkbook);
  const cleanedBytes = await removeLegacyCashflowColumnAn(exportedBytes);
  await fs.writeFile(destinationWorkbook, cleanedBytes);
}

async function removeLegacyCashflowColumnAn(xlsxBytes) {
  const archive = await JSZip.loadAsync(xlsxBytes);
  const workbookXml = await archive.file("xl/workbook.xml")?.async("string");
  const relationshipsXml = await archive.file("xl/_rels/workbook.xml.rels")?.async("string");
  if (!workbookXml || !relationshipsXml) fail("Exported workbook is missing its workbook metadata.");

  const cashflowSheetTag = [...workbookXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sheet\b[^>]*>/g)]
    .map((match) => match[0])
    .find((tag) => /\bname="CashFlow"/.test(tag));
  const relationshipId = cashflowSheetTag?.match(/\br:id="([^"]+)"/)?.[1];
  if (!relationshipId) fail("Exported workbook is missing the CashFlow worksheet relationship.");

  const relationshipTag = [...relationshipsXml.matchAll(/<Relationship\b[^>]*>/g)]
    .map((match) => match[0])
    .find((tag) => new RegExp(`\\bId="${relationshipId}"`).test(tag));
  const target = relationshipTag?.match(/\bTarget="([^"]+)"/)?.[1];
  if (!target) fail("Exported workbook is missing the CashFlow worksheet target.");

  const worksheetPath = target.startsWith("/")
    ? target.slice(1)
    : path.posix.normalize(path.posix.join("xl", target));
  const worksheetFile = archive.file(worksheetPath);
  if (!worksheetFile) fail(`Exported workbook is missing ${worksheetPath}.`);

  let worksheetXml = await worksheetFile.async("string");
  worksheetXml = worksheetXml.replace(/<(?:[A-Za-z_][\w.-]*:)?c\b(?=[^>]*\br="AN[1-9]\d*")[^>]*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?c>)/g, "");
  worksheetXml = worksheetXml.replace(/(<(?:[A-Za-z_][\w.-]*:)?dimension\b[^>]*\bref="[^"]*:)(AN)([1-9]\d*"[^>]*>)/, "$1AM$3");
  archive.file(worksheetPath, worksheetXml);
  return archive.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

function prepareCashflowOnlyPreview(workbook) {
  const cashflow = workbook.worksheets.getItem("CashFlow");
  for (const address of ["I12:J454", "V12:W456", "AJ12:AK500"]) {
    const range = cashflow.getRange(address);
    range.values = range.values;
  }

  for (const sheetName of listSheetNames(workbook).reverse()) {
    if (sheetName !== "CashFlow") workbook.worksheets.getItem(sheetName).delete();
  }
}

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

function inferCashflowFxRates(sheet) {
  const rates = new Map();
  for (const [currencyRange, rateRange] of [["G12:G454", "I12:I454"], ["T12:T456", "V12:V456"], ["AH12:AH500", "AJ12:AJ500"]]) {
    const currencies = sheet.getRange(currencyRange).values;
    const values = sheet.getRange(rateRange).values;
    for (let index = currencies.length - 1; index >= 0; index -= 1) {
      const currency = clean(currencies[index][0]).toUpperCase();
      const rate = Number(values[index][0]);
      if (currency && Number.isFinite(rate) && rate > 0 && !rates.has(currency)) rates.set(currency, rate);
    }
  }
  if (!rates.size) fail("CashFlow-only template does not contain usable FX rates.");
  return rates;
}

function refreshStandaloneTransferStatus(sheet) {
  const counterparties = sheet.getRange("AC12:AC500").values;
  const convertedAmounts = sheet.getRange("AK12:AK500").values;
  const totals = new Map();
  for (let index = 0; index < counterparties.length; index += 1) {
    const counterparty = clean(counterparties[index][0]);
    const amount = Number(convertedAmounts[index][0]);
    if (!counterparty || !Number.isFinite(amount)) continue;
    totals.set(counterparty, (totals.get(counterparty) || 0) + amount);
  }

  const seen = new Set();
  const statusRows = counterparties.map((row) => {
    const counterparty = clean(row[0]);
    if (!counterparty || seen.has(counterparty)) return [null, null];
    seen.add(counterparty);
    const outstanding = totals.get(counterparty) || 0;
    const status = outstanding === 0 ? "Settled" : outstanding < 0 ? "Owes You" : "You Owe";
    return [outstanding, status];
  });
  sheet.getRange("AL12:AM500").values = statusRows;
}

export function ensureCompactCashflowLayout(workbook) {
  const sheet = workbook.worksheets.getItem("CashFlow");
  const incomeHeaders = sheet.getRange("A11:L11").values[0].map(clean);
  const alreadyCompact = JSON.stringify(incomeHeaders.slice(0, 10)) === JSON.stringify(COMPACT_HEADERS.income);

  if (!alreadyCompact) {
    const legacyIncome = ["Date", "Year-Month", "Book", "Category", "Tag", "Description", "Amount", "Currency", "Account", "Notes", "FX Rate", "Converted Amount"];
    const legacyExpense = ["Date", "Year-Month", "Book", "Category", "Tag", "Description", "Amount", "Currency", "Account", "Note", "FX Rate", "Converted Amount"];
    const legacyTransfer = ["Date", "Year-Month", "Counterparty", "Category", "Tag", "Description", "Amount", "Currency", "Account", "Note", "FX Rate", "Converted Amount", "Outstanding", "Status"];
    assertHeaders(sheet, "A11:L11", legacyIncome);
    assertHeaders(sheet, "N11:Y11", legacyExpense);
    assertHeaders(sheet, "AA11:AN11", legacyTransfer);

    compactBlock(sheet, ["A", "B", "D", "E", "F", "G", "H", "I", "K", "L"], ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"], "K11:L500");
    compactBlock(sheet, ["N", "O", "Q", "R", "S", "T", "U", "V", "X", "Y"], ["N", "O", "P", "Q", "R", "S", "T", "U", "V", "W"], "X11:Y500");
    compactBlock(sheet, ["AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AK", "AL", "AM", "AN"], ["AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM"], "AN11:AN500");

    sheet.unmergeCells("A10:L10");
    sheet.unmergeCells("N10:Y10");
    sheet.unmergeCells("AA10:AN10");
    sheet.mergeCells("A10:J10");
    sheet.mergeCells("N10:W10");
    sheet.mergeCells("AA10:AM10");
    sheet.getRange("A11:J11").values = [COMPACT_HEADERS.income];
    sheet.getRange("N11:W11").values = [COMPACT_HEADERS.expense];
    sheet.getRange("AA11:AM11").values = [COMPACT_HEADERS.transfer];
    rewriteCashflowSummaryFormulas(sheet);
    rewriteWorkbookCashflowReferences(workbook);

    if (hasWorksheet(workbook, "Asset")) rewriteExistingCashflowFormulas(sheet);
  }

  sheet.getRange("A11:J11").values = [COMPACT_HEADERS.income];
  sheet.getRange("N11:W11").values = [COMPACT_HEADERS.expense];
  sheet.getRange("AA11:AM11").values = [COMPACT_HEADERS.transfer];
  rewriteCashflowSummaryFormulas(sheet);
  formatCashflowHeaders(sheet);
  formatExistingCashflowRows(sheet);
  sheet.getRange("K10:L10").clear({ applyTo: "all" });
  sheet.getRange("X10:Y10").clear({ applyTo: "all" });
  // Status now lives in AM. The exported XLSX is stripped of all legacy AN
  // cells by saveCashflowWorkbook, including formulas retained by importers.
  sheet.getRange("AN10:AN10").clear({ applyTo: "all" });
  if (!hasWorksheet(workbook, "Asset")) refreshStandaloneDerivedValues(sheet);
  return sheet;
}

function formatCashflowHeaders(sheet) {
  for (const address of ["A11:J11", "N11:W11", "AA11:AM11"]) {
    const range = sheet.getRange(address);
    range.format.fill = "#1F4E78";
    range.format.font.name = "Calibri";
    range.format.font.size = 11;
    range.format.font.bold = true;
    range.format.font.color = "#FFFFFF";
    range.format.borders = { bottom: { style: "thin", color: "#D9D9D9" } };
    range.format.wrapText = false;
    range.format.horizontalAlignment = "center";
    range.format.verticalAlignment = "center";
  }
}

function refreshStandaloneDerivedValues(sheet) {
  for (const [amountColumn, rateColumn, convertedColumn, startRow, endRow] of [
    ["F", "I", "J", 12, 454],
    ["S", "V", "W", 12, 456],
    ["AG", "AJ", "AK", 12, 500],
  ]) {
    const amounts = sheet.getRange(`${amountColumn}${startRow}:${amountColumn}${endRow}`).values;
    const rates = sheet.getRange(`${rateColumn}${startRow}:${rateColumn}${endRow}`).values;
    sheet.getRange(`${convertedColumn}${startRow}:${convertedColumn}${endRow}`).values = amounts.map((row, index) => {
      if (row[0] === null || row[0] === "") return [null];
      const amount = Number(row[0]);
      const rate = Number(rates[index][0]);
      return [Number.isFinite(amount) && Number.isFinite(rate) ? amount * rate : null];
    });
  }
  refreshStandaloneTransferStatus(sheet);
}

function compactBlock(sheet, sourceColumns, destinationColumns, tailRange) {
  const snapshots = sourceColumns.map((column) => sheet.getRange(`${column}11:${column}500`).values);
  for (let index = 0; index < sourceColumns.length; index += 1) {
    if (sourceColumns[index] === destinationColumns[index]) continue;
    sheet.getRange(`${destinationColumns[index]}11:${destinationColumns[index]}500`).copyFrom(
      sheet.getRange(`${sourceColumns[index]}11:${sourceColumns[index]}500`),
      "all",
    );
  }
  for (let index = 0; index < destinationColumns.length; index += 1) {
    sheet.getRange(`${destinationColumns[index]}11:${destinationColumns[index]}500`).values = snapshots[index];
  }
  sheet.getRange(tailRange).clear({ applyTo: "all" });
}

function rewriteCashflowSummaryFormulas(sheet) {
  sheet.getRange("A2").values = [["Selected Month："]];
  sheet.getRange("A3:A5").values = [["Total Income (Currency)"], ["Total Expense (Currency)"], ["Total Transfer (Currency)"]];
  sheet.getRange("D3:D5").values = [["Start Date"], ["Start Date"], ["Start Date"]];
  sheet.getRange("G3:G5").values = [["End Date"], ["End Date"], ["End Date"]];
  sheet.getRange("J3:J5").values = [["Income Record Count"], ["Expense Record Count"], ["Transfer Record Count"]];
  sheet.getRange("A7").values = [["Net Cash Flow"]];
  sheet.getRange("B3").formulas = [[`=SUMIFS($J$12:$J$454,$A$12:$A$454,">="&DATEVALUE($B$2&"-01"),$A$12:$A$454,"<="&EOMONTH(DATEVALUE($B$2&"-01"),0))`]];
  sheet.getRange("B4").formulas = [[`=SUMIFS($W$12:$W$456,$N$12:$N$456,">="&DATEVALUE($B$2&"-01"),$N$12:$N$456,"<="&EOMONTH(DATEVALUE($B$2&"-01"),0))`]];
  sheet.getRange("B5").formulas = [[`=SUMIFS($AK$12:$AK$500,$AB$12:$AB$500,$B$2)`]];
  sheet.getRange("E3").formulas = [[`=IF(K3=0,"",_xlfn.MINIFS($A$12:$A$454,$B$12:$B$454,$B$2))`]];
  sheet.getRange("H3").formulas = [[`=IF(K3=0,"",_xlfn.MAXIFS($A$12:$A$454,$B$12:$B$454,$B$2))`]];
  sheet.getRange("K3").formulas = [[`=COUNTIF($B$12:$B$454,$B$2)`]];
  sheet.getRange("E4").formulas = [[`=IF(K4=0,"",_xlfn.MINIFS($N$12:$N$456,$O$12:$O$456,$B$2))`]];
  sheet.getRange("H4").formulas = [[`=IF(K4=0,"",_xlfn.MAXIFS($N$12:$N$456,$O$12:$O$456,$B$2))`]];
  sheet.getRange("K4").formulas = [[`=COUNTIF($O$12:$O$456,$B$2)`]];
  sheet.getRange("E5").formulas = [[`=IF(K5=0,"",_xlfn.MINIFS($AA$12:$AA$500,$AB$12:$AB$500,$B$2))`]];
  sheet.getRange("H5").formulas = [[`=IF(K5=0,"",_xlfn.MAXIFS($AA$12:$AA$500,$AB$12:$AB$500,$B$2))`]];
  sheet.getRange("K5").formulas = [[`=COUNTIF($AB$12:$AB$500,$B$2)`]];
  sheet.getRange("B7").formulas = [[`=B3-B4`]];
}

function rewriteWorkbookCashflowReferences(workbook) {
  if (!hasWorksheet(workbook, "Summary")) return;
  const sheet = workbook.worksheets.getItem("Summary");
  const used = sheet.getUsedRange();
  const formulas = used.formulas;
  for (let row = 0; row < formulas.length; row += 1) {
    for (let column = 0; column < formulas[row].length; column += 1) {
      const formula = formulas[row][column];
      if (typeof formula !== "string" || !formula.includes("CashFlow!")) continue;
      const updated = formula.replaceAll("CashFlow!$L$", "CashFlow!$J$").replaceAll("CashFlow!$Y$", "CashFlow!$W$");
      if (updated !== formula) sheet.getCell(row, column).formulas = [[updated]];
    }
  }
}

function rewriteExistingCashflowFormulas(sheet) {
  for (const kind of ["income", "expense", "transfer"]) {
    const config = CASHFLOW_BLOCKS[kind];
    const dates = sheet.getRangeByIndexes(config.startRow - 1, config.startCol, config.endRow - config.startRow + 1, 1).values;
    for (let index = 0; index < dates.length; index += 1) {
      if (dates[index][0] === null || dates[index][0] === "") continue;
      writeCashflowFormulas(sheet, config.startRow + index, 1, kind, [], null);
    }
  }
}

function formatExistingCashflowRows(sheet) {
  for (const kind of ["income", "expense", "transfer"]) {
    const config = CASHFLOW_BLOCKS[kind];
    const dates = sheet.getRangeByIndexes(config.startRow - 1, config.startCol, config.endRow - config.startRow + 1, 1).values;
    let lastRow = config.startRow - 1;
    for (let index = 0; index < dates.length; index += 1) if (dates[index][0] !== null && dates[index][0] !== "") lastRow = config.startRow + index;
    if (lastRow >= config.startRow) formatCashRows(sheet, config.startRow, lastRow - config.startRow + 1, kind);
  }
}

function hasWorksheet(workbook, name) {
  return listSheetNames(workbook).includes(name);
}

function appendCashBlock(sheet, incoming, kind, fxRates = null) {
  const config = CASHFLOW_BLOCKS[kind];
  const existingRange = sheet.getRangeByIndexes(config.startRow - 1, config.startCol, config.endRow - config.startRow + 1, config.sourceWidth);
  const existing = existingRange.values;
  const counts = new Map();
  let lastRow = config.startRow - 1;

  for (let index = 0; index < existing.length; index += 1) {
    const row = existing[index];
    if (row[0] === null || row[0] === "") continue;
    lastRow = config.startRow + index;
    const dateText = excelDateText(row[0]);
    const key = kind === "income"
      ? keyOf([dateText, row[2], row[4], row[5], row[6], row[7]])
      : kind === "expense"
        ? keyOf([dateText, row[2], row[3], row[4], row[5], row[6], row[7]])
        : keyOf([dateText, row[2], row[3], row[5], row[6], row[7], row[8]]);
    counts.set(key, (counts.get(key) || 0) + 1);
  }

  const seen = new Map();
  const rowsToAdd = [];
  let skippedDuplicates = 0;
  for (const item of incoming) {
    const occurrence = (seen.get(item.key) || 0) + 1;
    seen.set(item.key, occurrence);
    if (occurrence <= (counts.get(item.key) || 0)) skippedDuplicates += 1;
    else rowsToAdd.push(item.values);
  }

  const firstRow = lastRow + 1;
  const requiredLastRow = firstRow + rowsToAdd.length - 1;
  if (requiredLastRow > config.endRow) fail(`${kind} block capacity exceeded; requires row ${requiredLastRow}, limit is ${config.endRow}.`);

  if (rowsToAdd.length) {
    sheet.getRangeByIndexes(firstRow - 1, config.startCol, rowsToAdd.length, config.sourceWidth).values = rowsToAdd;
    writeCashflowFormulas(sheet, firstRow, rowsToAdd.length, kind, rowsToAdd, fxRates);
    formatCashRows(sheet, firstRow, rowsToAdd.length, kind);
  }

  return {
    imported: rowsToAdd.length,
    skippedDuplicates,
    firstRow: rowsToAdd.length ? firstRow : null,
    lastRow: rowsToAdd.length ? requiredLastRow : null,
  };
}

function writeCashflowFormulas(sheet, firstRow, count, kind, sourceRows, fxRates) {
  const lastRow = firstRow + count - 1;
  if (fxRates) {
    const amountIndex = kind === "transfer" ? 6 : 5;
    const currencyIndex = kind === "transfer" ? 7 : 6;
    const converted = sourceRows.map((row) => {
      const currency = clean(row[currencyIndex]).toUpperCase();
      const rate = fxRates.get(currency);
      if (!rate) fail(`CashFlow-only template is missing an FX rate for ${currency}.`);
      return [rate, row[amountIndex] * rate];
    });
    const target = kind === "income" ? `I${firstRow}:J${lastRow}` : kind === "expense" ? `V${firstRow}:W${lastRow}` : `AJ${firstRow}:AK${lastRow}`;
    sheet.getRange(target).values = converted;
    return;
  }
  if (kind === "income") {
    sheet.getRange(`I${firstRow}:J${lastRow}`).formulas = Array.from({ length: count }, (_, index) => {
      const row = firstRow + index;
      return [
        `=IF(G${row}="","",IF(G${row}='Asset'!$K$3,1,IF(AND(G${row}="USD",'Asset'!$K$3="CNY"),'Asset'!$K$4,IF(AND(G${row}="CNY",'Asset'!$K$3="USD"),1/'Asset'!$K$4,""))))`,
        `=IF(OR(F${row}="",I${row}=""),"",F${row}*I${row})`,
      ];
    });
  } else if (kind === "expense") {
    sheet.getRange(`V${firstRow}:W${lastRow}`).formulas = Array.from({ length: count }, (_, index) => {
      const row = firstRow + index;
      return [
        `=IF(T${row}="","",IF(T${row}='Asset'!$K$3,1,IF(AND(T${row}="USD",'Asset'!$K$3="CNY"),'Asset'!$K$4,IF(AND(T${row}="CNY",'Asset'!$K$3="USD"),1/'Asset'!$K$4,""))))`,
        `=IF(OR(S${row}="",V${row}=""),"",S${row}*V${row})`,
      ];
    });
  } else {
    sheet.getRange(`AJ${firstRow}:AM${lastRow}`).formulas = Array.from({ length: count }, (_, index) => {
      const row = firstRow + index;
      return [
        `=IF(AH${row}="","",IF(AH${row}='Asset'!$K$3,1,IF(AND(AH${row}="USD",'Asset'!$K$3="CNY"),'Asset'!$K$4,IF(AND(AH${row}="CNY",'Asset'!$K$3="USD"),1/'Asset'!$K$4,""))))`,
        `=IF(OR(AG${row}="",AJ${row}=""),"",AG${row}*AJ${row})`,
        `=IF(COUNTIF($AC$12:AC${row},AC${row})>1,"",SUMIFS($AK$12:$AK$500,$AC$12:$AC$500,AC${row}))`,
        `=IF(AL${row}="","",IF(AL${row}=0,"Settled",IF(AL${row}<0,"Owes You","You Owe")))`,
      ];
    });
  }
}

function formatCashRows(sheet, firstRow, count, kind) {
  const lastRow = firstRow + count - 1;
  const columns = {
    income: { date: "A", month: "B", text: ["C", "D", "E", "G", "H"], amount: "F", rate: "I", converted: "J" },
    expense: { date: "N", month: "O", text: ["P", "Q", "R", "T", "U"], amount: "S", rate: "V", converted: "W" },
    transfer: { date: "AA", month: "AB", text: ["AC", "AD", "AE", "AF", "AH", "AI"], amount: "AG", rate: "AJ", converted: "AK", outstanding: "AL", status: "AM" },
  }[kind];
  const applyStyle = (range, { color, numberFormat }) => {
    range.format.font.name = "Calibri";
    range.format.font.size = 11;
    range.format.font.color = color;
    range.format.borders = { preset: "all", style: "thin", color: "#BFBFBF" };
    range.format.wrapText = false;
    range.format.horizontalAlignment = "center";
    range.format.verticalAlignment = "center";
    if (numberFormat) range.format.numberFormat = numberFormat;
  };

  applyStyle(sheet.getRange(`${columns.date}${firstRow}:${columns.date}${lastRow}`), { color: "#008000", numberFormat: "yyyy-mm-dd" });
  applyStyle(sheet.getRange(`${columns.month}${firstRow}:${columns.month}${lastRow}`), { color: "#008000" });
  for (const column of columns.text) {
    applyStyle(sheet.getRange(`${column}${firstRow}:${column}${lastRow}`), { color: "#008000" });
  }
  applyStyle(sheet.getRange(`${columns.amount}${firstRow}:${columns.amount}${lastRow}`), { color: "#008000", numberFormat: "#,##0.00;[Red](#,##0.00);-" });
  applyStyle(sheet.getRange(`${columns.rate}${firstRow}:${columns.rate}${lastRow}`), { color: "#008000", numberFormat: "0.0000" });
  applyStyle(sheet.getRange(`${columns.converted}${firstRow}:${columns.converted}${lastRow}`), { color: "#000000", numberFormat: "#,##0.00;[Red](#,##0.00);-" });
  if (columns.outstanding) applyStyle(sheet.getRange(`${columns.outstanding}${firstRow}:${columns.outstanding}${lastRow}`), { color: "#000000", numberFormat: "#,##0.00;[Red](#,##0.00);-" });
  if (columns.status) applyStyle(sheet.getRange(`${columns.status}${firstRow}:${columns.status}${lastRow}`), { color: "#000000" });
}

async function readCsvRecords(csvPath) {
  const text = (await fs.readFile(csvPath, "utf8")).replace(/^\uFEFF/, "");
  const csvBook = await Workbook.fromCSV(text, { sheetName: "CSV" });
  const values = csvBook.worksheets.getItem("CSV").getUsedRange(true).values;
  if (!values?.length) fail(`CSV is empty: ${csvPath}`);
  const headers = values[0].map(clean);
  const transferMarker = values.findIndex((row) => clean(row[0]) === "---转账记录---");
  const primaryRows = values.slice(1, transferMarker === -1 ? values.length : transferMarker);
  const records = primaryRows.filter((row) => row.some((value) => clean(value) !== "")).map((row, index) => {
    const record = { __row: index + 2 };
    headers.forEach((header, column) => { record[header] = row[column] ?? ""; });
    return record;
  });

  const transferRecords = [];
  if (transferMarker !== -1) {
    const transferHeaders = values[transferMarker + 1]?.slice(0, 5).map(clean) || [];
    const expected = ["日期", "转出账户", "转入账户", "备注", "金额"];
    if (JSON.stringify(transferHeaders) !== JSON.stringify(expected)) fail(`Unexpected transfer headers in ${csvPath}: ${transferHeaders.join(" | ")}`);
    for (let index = transferMarker + 2; index < values.length; index += 1) {
      const row = values[index];
      if (!row.slice(0, 5).some((value) => clean(value) !== "")) continue;
      const record = { __row: index + 1 };
      transferHeaders.forEach((header, column) => { record[header] = row[column] ?? ""; });
      transferRecords.push(record);
    }
  }
  return { headers, records, transferRecords };
}

function aggregateReports(reports) {
  const total = { sourceFiles: reports.length, sourceRows: 0, latestDate: null, excludedAfterCutoff: 0, income: { sourceRecords: 0, imported: 0, skippedDuplicates: 0 }, expense: { sourceRecords: 0, imported: 0, skippedDuplicates: 0 }, transfer: { sourceRecords: 0, imported: 0, skippedDuplicates: 0 } };
  for (const report of reports) {
    total.sourceRows += report.sourceRows;
    total.excludedAfterCutoff += report.excludedAfterCutoff.length;
    if (!total.latestDate || report.latestDate > total.latestDate) total.latestDate = report.latestDate;
    for (const kind of ["income", "expense", "transfer"]) {
      total[kind].sourceRecords += report[kind].sourceRecords;
      total[kind].imported += report[kind].imported;
      total[kind].skippedDuplicates += report[kind].skippedDuplicates;
    }
  }
  return total;
}

function parseCliArgs(argv) {
  const args = { cashflowCsvs: [] };
  for (let index = 0; index < argv.length; index += 1) {
    const key = argv[index];
    if (key === "--dry-run") args.dryRun = true;
    else if (key === "--allow-existing-output") args.allowExistingOutput = true;
    else if (key === "--cashflow-only") args.cashflowOnly = true;
    else if (key === "--cashflow-csv") args.cashflowCsvs.push(argv[++index]);
    else if (key.startsWith("--")) args[key.slice(2).replace(/-([a-z])/g, (_, character) => character.toUpperCase())] = argv[++index];
    else fail(`Unexpected argument: ${key}`);
  }
  return args;
}

function parseDate(value) {
  const datePart = clean(value).split(/[T\s]/)[0];
  const iso = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(datePart);
  const us = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(datePart);
  if (!iso && !us) return null;
  const [year, month, day] = iso ? [iso[1], iso[2], iso[3]] : [us[3], us[1], us[2]];
  const text = `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
  return { text, month: text.slice(0, 7), value: new Date(`${text}T00:00:00.000Z`) };
}

function parseNumber(value, label) {
  const original = String(value ?? "").trim();
  if (!original) fail(`Invalid ${label}: blank`);
  const negative = /^\(.*\)$/.test(original);
  const normalized = original.replace(/[()$,]/g, "").trim();
  const number = Number(normalized) * (negative ? -1 : 1);
  if (!Number.isFinite(number)) fail(`Invalid ${label}: ${value}`);
  return number;
}

function assertHeaders(sheet, range, expected) {
  const actual = sheet.getRange(range).values[0].map(clean);
  if (JSON.stringify(actual) !== JSON.stringify(expected)) fail(`Unexpected headers at ${sheet.name}!${range}: ${actual.join(" | ")}`);
}

function requireHeaders(actual, required) {
  for (const header of required) if (!actual.includes(header)) fail(`CSV missing required header: ${header}`);
}

function normalizeAccount(value) {
  const account = clean(value);
  if (["微信钱包", "微信", "Wechat", "WeChat"].includes(account)) return "WeChat";
  if (["Robinhood", "Robinhood Cash"].includes(account)) return "Robinhood Cash";
  if (["Fidelity", "Fidelity Cash"].includes(account)) return "Fidelity Cash";
  if (["IBKR", "IBKR Cash"].includes(account)) return "IBKR Cash";
  if (["Charles", "Charles Schwab", "Schwab", "Charles Cash"].includes(account)) return "Charles Cash";
  return account;
}

function excelDateText(value) {
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "number") return new Date(Date.UTC(1899, 11, 30) + Math.round(value) * 86400000).toISOString().slice(0, 10);
  return clean(value).slice(0, 10);
}

function keyOf(values) {
  return JSON.stringify(values.map((value) => typeof value === "string" ? value.trim() : value ?? null));
}

function clean(value) {
  return value === null || value === undefined ? "" : String(value).trim();
}

async function exists(file) {
  try {
    await fs.access(file);
    return true;
  } catch {
    return false;
  }
}

function fail(message) {
  throw new Error(message);
}

async function main() {
  const args = parseCliArgs(process.argv.slice(2));
  const inputDir = path.resolve(args.inputDir || process.cwd());
  const cashflowOnly = Boolean(args.cashflowOnly);
  let workbookPath = args.workbook || path.join(inputDir, "Output", "One_Piece.xlsx");
  if (cashflowOnly && !args.workbook) {
    const temporaryDirectory = await fs.mkdtemp(path.join(os.tmpdir(), "wealth-cashflow-preview-"));
    workbookPath = path.join(temporaryDirectory, "blank.xlsx");
    await createTreasuryWorkbook({ inputFile: path.join(inputDir, "monthly-close-input.md"), outputPath: workbookPath });
  }
  const cashflowDir = args.cashflowDir || path.join(inputDir, "CashFlow");
  if (!args.output) {
    fail("Usage: import_cashflow.mjs [--input-dir DIR] [--workbook FILE] --output FILE [--cashflow-dir DIR] [--cashflow-csv FILE ...] [--through-date YYYY-MM-DD] [--cashflow-only] [--dry-run] [--allow-existing-output]");
  }
  const report = await runCashflowImport({
    workbookPath,
    outputPath: args.output,
    cashflowDir,
    cashflowCsvs: args.cashflowCsvs,
    throughDate: args.throughDate || null,
    dryRun: Boolean(args.dryRun),
    allowExistingOutput: Boolean(args.allowExistingOutput),
    cashflowOnly,
  });
  console.log(JSON.stringify(report, null, 2));
}

const directRun = process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href;
if (directRun) main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exitCode = 1;
});

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

const ASSET_HEADERS = ["资产名称", "Category", "币种", "Amount", "FX Rate", "Converted Amount", "Bucket", "Update time"];
const DETAIL_FIRST_ROW = 13;
const DETAIL_LAST_ROW = 201;
const HOLDINGS_FIRST_ROW = 38;
const HOLDINGS_LAST_ROW = 67;
const PRICE_FIRST_ROW = 4;
const PRICE_LAST_ROW = 33;
const SECURITY_CATEGORIES = new Set(["Stock", "ETF"]);
const ACCOUNT_CATEGORIES = new Set(["Account", "Cash"]);
const KNOWN_ETFS = new Set(["DIA", "IWM", "JAAA", "JEPQ", "QQQ", "QQQI", "RSP", "SCHD", "SPMO", "SPY", "TQQQ", "VNQ", "VOO", "VTI"]);

export async function runAssetUpdate({
  workbookPath,
  outputPath,
  inputFile,
  assetDate = null,
  dryRun = false,
  allowExistingOutput = false,
}) {
  const sourceWorkbook = path.resolve(workbookPath);
  const destinationWorkbook = path.resolve(outputPath);
  const closeInput = path.resolve(inputFile);
  if (sourceWorkbook === destinationWorkbook) fail("Refusing to overwrite the source workbook. Choose a distinct output path.");
  if (!dryRun && !allowExistingOutput && await exists(destinationWorkbook)) fail(`Output already exists: ${destinationWorkbook}`);

  const snapshot = await readMonthlyCloseInput(closeInput, assetDate);
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(sourceWorkbook));
  const report = updateAssetSheet(workbook, snapshot);

  if (!dryRun) {
    await fs.mkdir(path.dirname(destinationWorkbook), { recursive: true });
    await saveWorkbookFormulaSafe(workbook, destinationWorkbook);
  }
  return {
    workbook: sourceWorkbook,
    output: dryRun ? null : destinationWorkbook,
    inputFile: closeInput,
    asOfDate: snapshot.date.text,
    ...report,
  };
}

export function updateAssetSheet(workbook, snapshot) {
  const asset = workbook.worksheets.getItem("Asset");
  const stock = workbook.worksheets.getItem("Stock");
  assertHeaders(asset, "A12:H12", ASSET_HEADERS);

  const detail = asset.getRange(`A${DETAIL_FIRST_ROW}:H${DETAIL_LAST_ROW}`).values;
  const accountRows = new Map();
  const securityRows = new Map();
  let lastPopulatedRow = DETAIL_FIRST_ROW - 1;
  for (let offset = 0; offset < detail.length; offset += 1) {
    const row = DETAIL_FIRST_ROW + offset;
    const [name, category, currency] = detail[offset];
    if (clean(name)) lastPopulatedRow = row;
    if (ACCOUNT_CATEGORIES.has(clean(category))) {
      const key = assetKey(name, currency);
      if (accountRows.has(key)) fail(`Duplicate asset account key in workbook: ${key}`);
      accountRows.set(key, { row, name: clean(name), currency: clean(currency).toUpperCase(), previousAmount: detail[offset][3] });
    } else if (SECURITY_CATEGORIES.has(clean(category))) {
      const ticker = normalizeTicker(name);
      if (securityRows.has(ticker)) fail(`Duplicate Asset security ticker: ${ticker}`);
      securityRows.set(ticker, { row, displayName: clean(name) });
    }
  }

  const accountChanges = [];
  const seen = new Set();
  for (const item of snapshot.balances) {
    const key = assetKey(item.account, item.currency);
    if (seen.has(key)) fail(`Duplicate balance in monthly close input: ${item.account} ${item.currency}`);
    seen.add(key);
    const match = accountRows.get(key);
    if (!match) fail(`Asset account not found in Asset!A13:C201: ${item.account} ${item.currency}.`);
    asset.getRange(`D${match.row}`).values = [[item.balance]];
    asset.getRange(`H${match.row}`).values = [[snapshot.date.value]];
    writeFxAndConvertedFormulas(asset, match.row);
    formatDetailRow(asset, match.row);
    accountChanges.push({
      row: match.row,
      asset: match.name,
      currency: match.currency,
      previousAmount: match.previousAmount,
      newAmount: item.balance,
    });
  }

  // Existing account/cash rows may contain missing legacy formulas. Normalize all
  // of them even when their balance was omitted and therefore remains unchanged.
  for (const { row } of accountRows.values()) {
    writeFxAndConvertedFormulas(asset, row);
    formatDetailRow(asset, row);
  }

  const prices = readStockPrices(stock);
  const holdings = readStockHoldings(stock);
  const addedSecurities = [];
  const linkedSecurities = [];
  for (const holding of holdings) {
    const price = prices.get(holding.ticker);
    if (!price) fail(`Stock current-price input is missing ${holding.ticker}.`);
    if (!["USD", "CNY"].includes(price.currency)) fail(`Unsupported Asset currency for ${holding.ticker}: ${price.currency || "blank"}`);
    let target = securityRows.get(holding.ticker);
    if (!target) {
      const row = firstBlankDetailRow(asset, Math.max(lastPopulatedRow + 1, DETAIL_FIRST_ROW));
      if (!row) fail(`Asset detail block has no blank row through ${DETAIL_LAST_ROW} for ${holding.ticker}.`);
      const templateRow = chooseSecurityTemplateRow(securityRows);
      asset.getRange(`A${row}:H${row}`).copyFrom(asset.getRange(`A${templateRow}:H${templateRow}`), "all");
      target = { row, displayName: holding.stockTicker };
      securityRows.set(holding.ticker, target);
      lastPopulatedRow = Math.max(lastPopulatedRow, row);
      addedSecurities.push({ ticker: holding.ticker, row });
    }

    const category = inferSecurityType(holding.ticker, holding.name);
    if (!category) fail(`Could not determine whether ${holding.ticker} is a Stock or ETF.`);
    asset.getRange(`A${target.row}:C${target.row}`).values = [[target.displayName, category, price.currency]];
    asset.getRange(`G${target.row}`).values = [[holding.bucket]];
    asset.getRange(`D${target.row}`).formulas = [[`='Stock'!V${holding.row}`]];
    writeFxAndConvertedFormulas(asset, target.row);
    asset.getRange(`H${target.row}`).formulas = [[`='Stock'!S${price.row}`]];
    formatDetailRow(asset, target.row);
    linkedSecurities.push({ ticker: holding.ticker, assetRow: target.row, stockHoldingRow: holding.row, stockPriceRow: price.row });
  }

  return {
    updatedBalances: accountChanges.length,
    accountChanges,
    linkedSecurities,
    addedSecurities,
    omittedBalances: [...accountRows.values()].filter((item) => !seen.has(assetKey(item.name, item.currency))).map((item) => ({ asset: item.name, currency: item.currency, row: item.row })),
  };
}

async function readMonthlyCloseInput(inputFile, fallbackDate) {
  const text = await fs.readFile(inputFile, "utf8");
  const embeddedDate = text.match(/^As-of Date:\s*(.*?)\s*$/mi)?.[1] || "";
  const date = parseDate(fallbackDate || embeddedDate);
  if (!date) fail("Asset balances require an As-of Date in monthly-close-input.md or --asset-date YYYY-MM-DD.");

  const lines = text.split(/\r?\n/);
  const headerIndex = lines.findIndex((line) => /^\s*\|\s*Account\s*\|\s*Currency\s*\|\s*Balance\s*\|\s*$/i.test(line));
  if (headerIndex < 0) fail(`Could not find the Account | Currency | Balance table in ${inputFile}.`);
  const balances = [];
  for (let index = headerIndex + 2; index < lines.length; index += 1) {
    const line = lines[index].trim();
    if (!line.startsWith("|")) break;
    const cells = line.slice(1, line.endsWith("|") ? -1 : undefined).split("|").map(clean);
    if (cells.length < 3) continue;
    const account = normalizeAccount(cells[0]);
    const currency = cells[1].toUpperCase();
    if (!cells[2]) continue;
    if (!account || !["USD", "CNY"].includes(currency)) fail(`Invalid Asset input row ${index + 1}: ${line}`);
    balances.push({ account, currency, balance: parseNumber(cells[2], `balance at row ${index + 1}`), sourceRow: index + 1 });
  }
  if (!balances.length) fail(`No nonblank Asset balances found in ${inputFile}.`);
  return { date, balances };
}

function readStockPrices(sheet) {
  const values = sheet.getRange(`P${PRICE_FIRST_ROW}:S${PRICE_LAST_ROW}`).values;
  const prices = new Map();
  for (let offset = 0; offset < values.length; offset += 1) {
    const ticker = normalizeTicker(values[offset][0]);
    if (!ticker) continue;
    if (prices.has(ticker)) fail(`Duplicate Stock current-price ticker: ${ticker}`);
    prices.set(ticker, { ticker, currency: clean(values[offset][2]).toUpperCase(), row: PRICE_FIRST_ROW + offset });
  }
  return prices;
}

function readStockHoldings(sheet) {
  const values = sheet.getRange(`P${HOLDINGS_FIRST_ROW}:Z${HOLDINGS_LAST_ROW}`).values;
  const holdings = [];
  const seen = new Set();
  for (let offset = 0; offset < values.length; offset += 1) {
    const ticker = normalizeTicker(values[offset][0]);
    if (!ticker) continue;
    if (seen.has(ticker)) fail(`Duplicate Stock holdings ticker: ${ticker}`);
    seen.add(ticker);
    const bucket = clean(values[offset][10]);
    if (!bucket) fail(`Stock holdings Bucket is blank for ${ticker} at row ${HOLDINGS_FIRST_ROW + offset}.`);
    holdings.push({ ticker, stockTicker: clean(values[offset][0]), name: clean(values[offset][1]), bucket, row: HOLDINGS_FIRST_ROW + offset });
  }
  if (!holdings.length) fail(`Stock holdings summary has no tickers at Stock!P${HOLDINGS_FIRST_ROW}:P${HOLDINGS_LAST_ROW}.`);
  return holdings;
}

function writeFxAndConvertedFormulas(sheet, row) {
  sheet.getRange(`E${row}`).formulas = [[`=IF(C${row}="","",IF(C${row}=$K$3,1,IF(AND(C${row}="USD",$K$3="CNY"),$K$4,IF(AND(C${row}="CNY",$K$3="USD"),1/$K$4,""))))`]];
  sheet.getRange(`F${row}`).formulas = [[`=IF(OR(D${row}="",E${row}=""),"",D${row}*E${row})`]];
}

function formatDetailRow(sheet, row) {
  sheet.getRange(`D${row}`).format.numberFormat = "#,##0.00;[Red]\\(#,##0.00\\);\\-";
  sheet.getRange(`E${row}`).format.numberFormat = "0.0000";
  sheet.getRange(`F${row}`).format.numberFormat = "#,##0.00;[Red]\\(#,##0.00\\);\\-";
  sheet.getRange(`H${row}`).format.numberFormat = "yyyy-mm-dd";
}

function chooseSecurityTemplateRow(securityRows) {
  const rows = [...securityRows.values()].map((item) => item.row);
  return rows.length ? Math.max(...rows) : 22;
}

function firstBlankDetailRow(sheet, startRow) {
  for (let row = startRow; row <= DETAIL_LAST_ROW; row += 1) if (!clean(sheet.getRange(`A${row}`).values[0][0])) return row;
  for (let row = DETAIL_FIRST_ROW; row < startRow; row += 1) if (!clean(sheet.getRange(`A${row}`).values[0][0])) return row;
  return null;
}

function inferSecurityType(ticker, name) {
  if (KNOWN_ETFS.has(ticker) || /\bETF\b|EXCHANGE[- ]TRADED|INDEX FUND/i.test(name)) return "ETF";
  if (/^[A-Z][A-Z0-9.\/-]{0,9}$/.test(ticker) && name) return "Stock";
  return null;
}

function normalizeTicker(value) {
  const ticker = clean(value).toUpperCase();
  if (ticker === "TESLA") return "TSLA";
  if (ticker === "ECHO") return "SATS";
  return ticker;
}

function normalizeAccount(value) {
  const account = clean(value);
  if (["微信钱包", "微信", "Wechat", "WeChat"].includes(account)) return "WeChat";
  if (["Robinhood", "Robinhood Cash"].includes(account)) return "Robinhood Cash";
  if (["Fidelity", "Fidelity Cash"].includes(account)) return "Fidelity Cash";
  if (["IBKR", "IBKR Cash"].includes(account)) return "IBKR Cash";
  return account;
}

function assetKey(name, currency) { return `${normalizeAccount(name).toLowerCase()}|${clean(currency).toUpperCase()}`; }

function assertHeaders(sheet, range, expected) {
  const actual = sheet.getRange(range).values[0].map(clean);
  if (JSON.stringify(actual) !== JSON.stringify(expected)) fail(`Unexpected headers at ${sheet.name}!${range}: ${actual.join(" | ")}`);
}

function parseDate(value) {
  const text = clean(value);
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text);
  if (!match) return null;
  const date = new Date(`${text}T00:00:00.000Z`);
  if (Number.isNaN(date.getTime()) || date.toISOString().slice(0, 10) !== text) return null;
  return { text, value: date };
}

function parseNumber(value, label) {
  const original = clean(value);
  const parenthesized = /^\(.*\)$/.test(original);
  const number = Number(original.replace(/[()$,]/g, "").trim()) * (parenthesized ? -1 : 1);
  if (!original || !Number.isFinite(number)) fail(`Invalid ${label}: ${value}`);
  return number;
}

function clean(value) { return value === null || value === undefined ? "" : String(value).trim(); }
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
  const inputFile = path.resolve(args.inputFile || path.join(inputDir, "monthly-close-input.md"));
  if (!args.output) fail("Usage: update_asset.mjs [--input-dir DIR] [--workbook FILE] --output FILE [--input-file FILE] [--asset-date YYYY-MM-DD] [--dry-run] [--allow-existing-output]");
  const report = await runAssetUpdate({
    workbookPath,
    outputPath: args.output,
    inputFile,
    assetDate: args.assetDate || null,
    dryRun: Boolean(args.dryRun),
    allowExistingOutput: Boolean(args.allowExistingOutput),
  });
  console.log(JSON.stringify(report, null, 2));
}

const directRun = process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href;
if (directRun) main().catch((error) => { console.error(error instanceof Error ? error.message : String(error)); process.exitCode = 1; });

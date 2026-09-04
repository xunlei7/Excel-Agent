#!/usr/bin/env node

import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { createHash } from "node:crypto";
import { execFile } from "node:child_process";
import { createRequire } from "node:module";
import { fileURLToPath, pathToFileURL } from "node:url";
import { promisify } from "node:util";
import { saveCashflowWorkbook } from "./import_cashflow.mjs";
import { readBootstrap } from "./create_workbook.mjs";

const dependencies = process.env.CODEX_NODE_MODULES;
if (!dependencies) throw new Error("CODEX_NODE_MODULES is required");
const require = createRequire(path.join(dependencies, "package.json"));
const { FileBlob, SpreadsheetFile, Workbook } = require("@oai/artifact-tool");
const execFileAsync = promisify(execFile);
const STATEMENT_CACHE_VERSION = 1;
const STATEMENT_PARSER_VERSION = "2026-09-02.1";

const BROKER_FOLDERS = {
  Fidelity: { broker: "Fidelity", currency: "USD" },
  Robinhood: { broker: "Robinhood", currency: "USD" },
  Charles: { broker: "Charles", currency: "USD" },
  IBKR: { broker: "IBKR", currency: "USD" },
};
const TRANSACTION_HEADERS = ["Broker", "Date", "Ticker", "Name", "Action", "Units", "Trade Price", "Trade Value", "Current Price", "Current Value", "未实现盈亏", "币种", "Fee", "Cost Basis", "已实现盈亏"];
const HOLDINGS_HEADERS = ["Ticker", "名称", "Total Units", "Total Cost", "Avg Cost", "Current Price", "Current Value", "未实现盈亏", "已实现盈亏", "总收益", "Bucket"];
const PRICE_HEADERS = ["Ticker", "Current Price", "币种", "更新日期"];
const PRICE_FIRST_ROW = 4;
const PRICE_LAST_ROW = 33;
const HOLDINGS_HEADER_ROW = 37;
const HOLDINGS_FIRST_ROW = 38;
const HOLDINGS_LAST_ROW = 67;
const KNOWN_ETFS = new Set(["DIA", "IWM", "JAAA", "JEPQ", "QQQ", "QQQI", "RSP", "SCHD", "SPMO", "SPY", "TQQQ", "VNQ", "VOO", "VTI"]);
const SECURITY_NAMES = new Map([
  ["AAPL", "Apple Inc."],
  ["AMZN", "Amazon.com, Inc."],
  ["ARM", "Arm Holdings plc"],
  ["GOOG", "Alphabet Inc. Class C"],
  ["IONQ", "IonQ, Inc."],
  ["MSFT", "Microsoft Corporation"],
  ["MSTR", "Strategy Inc."],
  ["NVDA", "NVIDIA Corporation"],
  ["QQQ", "Invesco QQQ Trust"],
  ["RGTI", "Rigetti Computing, Inc."],
  ["SATS", "EchoStar Corporation"],
  ["TSLA", "Tesla, Inc."],
  ["VOO", "Vanguard S&P 500 ETF"],
  ["JAAA", "Janus Henderson AAA CLO ETF"],
  ["JEPQ", "JPMorgan Nasdaq Equity Premium Income ETF"],
  ["QQQI", "NEOS Nasdaq-100 High Income ETF"],
  ["RSP", "Invesco S&P 500 Equal Weight ETF"],
  ["SCHD", "Schwab U.S. Dividend Equity ETF"],
  ["SPMO", "Invesco S&P 500 Momentum ETF"],
  ["TQQQ", "ProShares UltraPro QQQ"],
  ["VNQ", "Vanguard Real Estate ETF"],
]);
const FIDELITY_SECURITY_IDS = new Map([
  ["594918104", "MSFT"],
  ["922908363", "VOO"],
]);

export async function discoverStockSources(inputDir, { allowEmpty = false } = {}) {
  const root = path.resolve(inputDir);
  const sources = [];
  const ignored = [];
  for (const [folder, defaults] of Object.entries(BROKER_FOLDERS)) {
    const directory = path.join(root, folder);
    if (!await exists(directory)) continue;
    const entries = (await fs.readdir(directory, { withFileTypes: true })).sort((a, b) => a.name.localeCompare(b.name));
    for (const entry of entries) {
      if (!entry.isFile() || entry.name.startsWith("~$")) continue;
      const file = path.join(directory, entry.name);
      const extension = path.extname(entry.name).toLowerCase();
      if (extension === ".csv") {
        const headers = await readCsvHeaders(file);
        if (looksLikeStockCsv(headers, defaults)) sources.push({ file, format: "csv", defaultBroker: defaults.broker, defaultCurrency: defaults.currency });
        else ignored.push({ file, reason: "unrecognized brokerage CSV headers" });
      } else if (extension === ".pdf" && folder === "Fidelity") {
        sources.push({ file, format: "fidelity-pdf", defaultBroker: defaults.broker, defaultCurrency: defaults.currency });
      } else {
        ignored.push({ file, reason: extension === ".pdf" ? `PDF profile is not yet defined for ${defaults.broker}` : "unsupported file type" });
      }
    }
  }
  if (!allowEmpty && !sources.length) fail(`No supported brokerage statements found under ${root}.`);
  return { root, sources, ignored };
}

export async function importStock(workbook, source, { allowNewStockRows = false, cacheDir = null, useCache = null, refreshCache = false } = {}) {
  const normalizedSource = typeof source === "string" ? inferSource(source) : { ...source };
  const cacheEnabled = useCache ?? Boolean(cacheDir);
  const { parsed, cache } = await readOrParseStockSource(normalizedSource, { cacheDir, useCache: cacheEnabled, refreshCache });
  return { ...applyStockTransactions(workbook, normalizedSource, parsed, { allowNewStockRows }), cache };
}

export async function runStockImport({ workbookPath, outputPath, inputDir, inputFile = null, stockSources = [], dryRun = false, allowExistingOutput = false, stockOnly = false, allowNewStockRows = false, fetchPrices = true, priceAsOf = null, cacheDir = null, useCache = true, refreshCache = false }) {
  const sourceWorkbook = path.resolve(workbookPath);
  const destinationWorkbook = path.resolve(outputPath);
  const resolvedCacheDir = path.resolve(cacheDir || path.join(inputDir, ".cache", "update-treasury", "stock-statements"));
  if (sourceWorkbook === destinationWorkbook) fail("Refusing to overwrite the source workbook. Choose a distinct output path.");
  if (!dryRun && !allowExistingOutput && await exists(destinationWorkbook)) fail(`Output already exists: ${destinationWorkbook}`);

  const discovery = stockSources.length ? null : await discoverStockSources(inputDir);
  const sources = stockSources.length ? stockSources.map(inferSource) : discovery.sources;
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(sourceWorkbook));
  const bootstrap = await readBootstrap(path.resolve(inputFile || path.join(inputDir, "monthly-close-input.md")));
  const stockSheet = workbook.worksheets.getItem("Stock");
  const openingSecurities = ensureStockRows(
    stockSheet,
    bootstrap.openingHoldings.map((item) => ({
      action: "Buy",
      ticker: item.ticker,
      name: canonicalSecurityName(item.ticker, item.ticker),
      currency: "USD",
    })),
    readTickerMetadata(stockSheet),
    { allowNewStockRows: true },
  );
  const reports = [];
  for (const source of sources) reports.push(await importStock(workbook, source, {
    allowNewStockRows: stockOnly || allowNewStockRows,
    cacheDir: resolvedCacheDir,
    useCache,
    refreshCache,
  }));
  normalizeSecurityNames(stockSheet);
  const transactionSort = sortTransactionLedger(stockSheet);
  const inferredCostBasis = backfillClosedPositionCostBasis(stockSheet);
  rewriteExistingTransactionFormulas(stockSheet);
  const priceUpdate = fetchPrices ? await fetchAndApplyCurrentPrices(stockSheet, inputDir, priceAsOf) : null;
  refreshHoldingsSummary(stockSheet, bootstrap);
  formatStockRows(stockSheet, 21, transactionSort.rowCount);
  formatStockReferenceBlocks(stockSheet);
  const finalMissingCurrentPrices = holdingsRows(stockSheet).filter(({ ticker }) => clean(readPriceRows(stockSheet).get(ticker)?.price) === "").map(({ ticker }) => ticker);
  if (stockOnly) prepareStockOnlyPreview(workbook);

  if (!dryRun) {
    await fs.mkdir(path.dirname(destinationWorkbook), { recursive: true });
    if (stockOnly) {
      const output = await SpreadsheetFile.exportXlsx(workbook);
      await output.save(destinationWorkbook);
    } else {
      await saveCashflowWorkbook(workbook, destinationWorkbook);
    }
  }
  return {
    workbook: sourceWorkbook,
    output: dryRun ? null : destinationWorkbook,
    recognizedFiles: sources.map((item) => item.file),
    ignoredFiles: discovery?.ignored || [],
    stockOnly,
    allowNewStockRows: stockOnly || allowNewStockRows,
    statementCache: summarizeStatementCache(reports, resolvedCacheDir, useCache),
    openingDate: bootstrap.openingDateText,
    openingHoldings: bootstrap.openingHoldings,
    openingSecurities,
    transactionSort,
    inferredCostBasis,
    priceUpdate,
    finalMissingCurrentPrices,
    files: reports,
    totals: summarizeReports(reports),
  };
}

async function readOrParseStockSource(source, { cacheDir, useCache, refreshCache }) {
  if (!useCache) return { parsed: await parseStockSource(source), cache: { status: "disabled", file: null } };
  if (!cacheDir) fail("Statement cache directory is required when caching is enabled.");

  const sourceBytes = await fs.readFile(source.file);
  const sourceHash = createHash("sha256").update(sourceBytes).digest("hex");
  const cacheKey = createHash("sha256")
    .update(JSON.stringify({ sourceHash, format: source.format, broker: source.defaultBroker, currency: source.defaultCurrency, parser: STATEMENT_PARSER_VERSION }))
    .digest("hex");
  const cacheFile = path.join(cacheDir, `${cacheKey}.json`);
  let invalidCache = false;

  if (!refreshCache && await exists(cacheFile)) {
    try {
      const cached = JSON.parse(await fs.readFile(cacheFile, "utf8"));
      return {
        parsed: hydrateCachedParse(cached, sourceHash, source),
        cache: { status: "hit", file: cacheFile, sourceHash },
      };
    } catch {
      invalidCache = true;
    }
  }

  const parsed = await parseStockSource(source);
  const payload = {
    cacheVersion: STATEMENT_CACHE_VERSION,
    parserVersion: STATEMENT_PARSER_VERSION,
    sourceHash,
    format: source.format,
    broker: source.defaultBroker,
    currency: source.defaultCurrency,
    parsed: serializeParsedStockSource(parsed),
  };
  await fs.mkdir(cacheDir, { recursive: true });
  const temporaryFile = `${cacheFile}.${process.pid}.tmp`;
  await fs.writeFile(temporaryFile, `${JSON.stringify(payload)}\n`, "utf8");
  await fs.rename(temporaryFile, cacheFile);
  return {
    parsed,
    cache: { status: refreshCache ? "refreshed" : invalidCache ? "repaired" : "miss", file: cacheFile, sourceHash },
  };
}

async function parseStockSource(source) {
  return source.format === "fidelity-pdf" ? parseFidelityStatement(source) : parseBrokerCsv(source);
}

function serializeParsedStockSource(parsed) {
  return {
    ...parsed,
    transactions: parsed.transactions.map((item) => ({ ...item, date: item.date.text })),
  };
}

function hydrateCachedParse(cached, sourceHash, source) {
  if (cached?.cacheVersion !== STATEMENT_CACHE_VERSION
    || cached?.parserVersion !== STATEMENT_PARSER_VERSION
    || cached?.sourceHash !== sourceHash
    || cached?.format !== source.format
    || cached?.broker !== source.defaultBroker
    || cached?.currency !== source.defaultCurrency
    || !Array.isArray(cached?.parsed?.transactions)
    || !Array.isArray(cached?.parsed?.ignored)) throw new Error("Invalid statement cache entry.");
  return {
    ...cached.parsed,
    transactions: cached.parsed.transactions.map((item) => {
      const date = parseDate(item.date);
      if (!date) throw new Error("Invalid cached transaction date.");
      return { ...item, date };
    }),
  };
}

async function parseBrokerCsv(source) {
  const rows = await readCsvRecords(source.file);
  const aliases = {
    broker: ["Broker", "券商", "经纪商"],
    date: ["Date", "Trade Date", "Activity Date", "Date/Time", "日期"],
    ticker: ["Ticker", "Symbol", "Instrument", "代码", "股票代码"],
    name: ["Name", "Security", "Description", "证券名称", "名称"],
    action: ["Action", "Side", "Type", "Trans Code", "Buy/Sell", "操作", "买卖"],
    units: ["Units", "Quantity", "Qty", "数量", "股数"],
    price: ["Trade Price", "Price", "T. Price", "成交价", "价格"],
    currency: ["Currency", "币种", "货币"],
    fee: ["Fee", "Fees & Comm", "Fees & Commissions", "Comm/Fee", "Commission", "费用", "佣金"],
    basis: ["Cost Basis", "Basis", "成本基础", "成本"],
    amount: ["Amount", "Net Amount", "Cash Amount", "Proceeds", "现金金额"],
  };
  const columns = Object.fromEntries(Object.entries(aliases).map(([key, names]) => [key, names.find((name) => rows.headers.includes(name))]));
  for (const required of ["date", "ticker", "action"]) if (!columns[required]) fail(`${source.file} is missing a recognized ${required} column.`);
  if (!source.defaultBroker) fail(`${source.file} must be stored in a recognized broker folder.`);
  if (!columns.currency && !source.defaultCurrency) fail(`${source.file} needs a Currency column or a broker-folder default.`);

  const transactions = [];
  const ignored = [];
  for (const raw of rows.records) {
    const rawAction = clean(raw[columns.action]);
    const rawTicker = clean(raw[columns.ticker]);
    const rawDate = clean(raw[columns.date]);
    if (!rawAction && !rawTicker && !rawDate) {
      ignored.push({ row: raw.__row, code: "", reason: "blank or footer row" });
      continue;
    }
    const action = canonicalAction(rawAction);
    if (!action) {
      if (isCashActivity(rawAction, columns.name ? raw[columns.name] : "")) {
        ignored.push({ row: raw.__row, code: rawAction, reason: "non-security cash activity" });
        continue;
      }
      fail(`Unsupported stock action at ${source.file} row ${raw.__row}: ${rawAction}`);
    }
    const ticker = normalizeTicker(rawTicker);
    if (!ticker) fail(`Missing ticker at ${source.file} row ${raw.__row}.`);
    const date = parseDate(rawDate);
    if (!date) fail(`Invalid stock date at ${source.file} row ${raw.__row}: ${rawDate}`);
    const broker = source.defaultBroker;
    const currency = (columns.currency ? clean(raw[columns.currency]).toUpperCase() : "") || source.defaultCurrency;
    if (!["USD", "CNY"].includes(currency)) fail(`Unsupported stock currency at ${source.file} row ${raw.__row}: ${currency}`);
    const units = action === "Dividend" ? optionalNumber(columns.units ? raw[columns.units] : "", 0) : parseNumber(raw[columns.units], "units");
    const price = action === "Dividend" ? optionalNumber(columns.price ? raw[columns.price] : "", 0) : parseNumber(raw[columns.price], "trade price");
    const fee = columns.fee && clean(raw[columns.fee]) !== "" ? Math.abs(parseNumber(raw[columns.fee], "fee")) : null;
    const basis = columns.basis && clean(raw[columns.basis]) !== "" ? Math.abs(parseNumber(raw[columns.basis], "cost basis")) : null;
    const dividendAmount = action === "Dividend" ? Math.abs(parseNumber(columns.amount ? raw[columns.amount] : "", "dividend amount")) : null;
    const name = columns.name ? clean(raw[columns.name]).split(/\r?\n/)[0] : "";
    transactions.push({ broker, date, ticker, name, action, units: Math.abs(units), price: Math.abs(price), currency, fee, basis, dividendAmount, sourceRow: raw.__row });
  }
  return { sourceRows: rows.records.length, transactions: aggregateTransactions(transactions), ignored, dateBasis: "source trade/activity date" };
}

async function parseFidelityStatement(source) {
  const { getDocument } = await import(pathToFileURL(require.resolve("pdfjs-dist/legacy/build/pdf.mjs")).href);
  const document = await getDocument({ data: new Uint8Array(await fs.readFile(source.file)), disableWorker: true }).promise;
  const pages = [];
  const allTokens = [];
  for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber += 1) {
    const page = await document.getPage(pageNumber);
    const content = await page.getTextContent();
    const tokens = content.items.map((item) => clean(item.str)).filter(Boolean);
    allTokens.push(...tokens);
    pages.push({ pageNumber, lines: groupPdfLines(content.items) });
  }
  const period = allTokens.find((value) => /[A-Za-z]+ \d{1,2}, \d{4}\s*-\s*[A-Za-z]+ \d{1,2}, \d{4}/.test(value));
  const year = period?.match(/-\s*[A-Za-z]+ \d{1,2}, (\d{4})/)?.[1];
  if (!year) fail(`Could not determine the statement year in ${source.file}.`);

  const nameToTicker = new Map();
  for (let index = 1; index < allTokens.length; index += 1) {
    const tickerMatch = /^\(([A-Za-z][A-Za-z0-9.\/-]{0,9})\)$/.exec(allTokens[index]);
    if (!tickerMatch) continue;
    let previous = index - 1;
    while (previous >= 0 && ["USD", "S"].includes(allTokens[previous])) previous -= 1;
    if (previous >= 0) nameToTicker.set(normalizeSecurityName(allTokens[previous]), normalizeTicker(tickerMatch[1]));
  }

  const transactions = [];
  const ignored = [];
  for (const page of pages) {
    for (const line of page.lines) {
      const tokens = line.tokens.map(clean).filter(Boolean);
      if (!/^\d{2}\/\d{2}$/.test(tokens[0] || "")) continue;
      const actionIndex = tokens.findIndex((value) => /^(You Bought|You Sold|Dividend Received)$/i.test(value));
      if (actionIndex < 0 || actionIndex < 3) continue;
      const [month, day] = tokens[0].split("/");
      const date = parseDate(`${year}-${month}-${day}`);
      const securityName = tokens[1];
      const action = canonicalAction(tokens[actionIndex]);
      const securityId = clean(tokens[2]);
      const ticker = nameToTicker.get(normalizeSecurityName(securityName)) || FIDELITY_SECURITY_IDS.get(securityId);
      if (!ticker || ticker === "SPAXX") {
        ignored.push({ page: page.pageNumber, date: date.text, security: securityName, reason: ticker === "SPAXX" || /MONEY MARKET/i.test(securityName) ? "core money-market income" : "security ticker not found in statement holdings" });
        continue;
      }
      if (action === "Dividend") {
        transactions.push({ broker: source.defaultBroker, date, ticker, name: securityName, action, units: 0, price: 0, currency: source.defaultCurrency, fee: null, basis: null, dividendAmount: Math.abs(parseNumber(tokens.at(-1), "dividend amount")), sourceRow: `page ${page.pageNumber}` });
      } else {
        const tradeDate = previousWeekday(date);
        const rawBasis = clean(tokens[actionIndex + 3]);
        const rawFee = clean(tokens.at(-2));
        const basis = action === "Sell" && rawBasis && rawBasis !== "-" ? Math.abs(parseNumber(rawBasis, "cost basis")) : null;
        const fee = action === "Sell" && rawFee && rawFee !== "-" && rawFee.toLowerCase() !== "f" ? Math.abs(parseNumber(rawFee, "fee")) : null;
        transactions.push({ broker: source.defaultBroker, date: tradeDate, ticker, name: securityName, action, units: Math.abs(parseNumber(tokens[actionIndex + 1], "quantity")), price: Math.abs(parseNumber(tokens[actionIndex + 2], "trade price")), currency: source.defaultCurrency, fee, basis, dividendAmount: null, sourceRow: `page ${page.pageNumber}` });
      }
    }
  }
  return { sourceRows: transactions.length + ignored.length, transactions: aggregateTransactions(transactions), ignored, dateBasis: "dividend date; trade date inferred as the previous weekday from Fidelity settlement date" };
}

function applyStockTransactions(workbook, source, parsed, { allowNewStockRows }) {
  const sheet = workbook.worksheets.getItem("Stock");
  assertHeaders(sheet, "A20:O20", TRANSACTION_HEADERS);
  assertHeaders(sheet, "P3:S3", PRICE_HEADERS);
  assertHeaders(sheet, `P${HOLDINGS_HEADER_ROW}:Z${HOLDINGS_HEADER_ROW}`, HOLDINGS_HEADERS);
  for (const item of parsed.transactions) item.name = canonicalSecurityName(item.ticker, item.name);
  const metadata = readTickerMetadata(sheet);
  const newSecurities = ensureStockRows(sheet, parsed.transactions, metadata, { allowNewStockRows });

  for (const item of parsed.transactions) {
    const known = metadata.get(item.ticker);
    item.name = canonicalSecurityName(item.ticker, known?.name || item.name);
  }
  const existing = sheet.getRange("A21:O212").values;
  const coveredTransactionGroups = findCoveredTransactionGroups(existing, parsed.transactions);
  const usedExisting = new Set();
  const add = [];
  const dividendBackfills = [];
  let lastRow = 20;
  for (let index = 0; index < existing.length; index += 1) if (clean(existing[index][0])) lastRow = index + 21;

  for (const item of parsed.transactions.sort((a, b) => a.date.text.localeCompare(b.date.text))) {
    if (item.action !== "Dividend" && coveredTransactionGroups.has(transactionGroupKey(item))) continue;
    let matchedIndex = -1;
    for (let index = 0; index < existing.length; index += 1) {
      if (usedExisting.has(index) || !clean(existing[index][0])) continue;
      if (sameTransaction(existing[index], item)) { matchedIndex = index; break; }
    }
    if (matchedIndex >= 0) {
      usedExisting.add(matchedIndex);
      if (item.action === "Dividend") {
        const row = matchedIndex + 21;
        const previous = Number(existing[matchedIndex][14]) || 0;
        if (Math.abs(previous - item.dividendAmount) > 0.000001) {
          sheet.getRange(`O${row}`).values = [[item.dividendAmount]];
          dividendBackfills.push({ row, previous, amount: item.dividendAmount });
        }
      }
      continue;
    }
    add.push(item);
  }

  const firstRow = lastRow + 1;
  const requiredLastRow = firstRow + add.length - 1;
  if (requiredLastRow > 212) fail(`Stock transaction block capacity exceeded; requires row ${requiredLastRow}, limit is 212.`);
  const templateRow = Math.max(21, lastRow);
  for (let index = 0; index < add.length; index += 1) {
    const row = firstRow + index;
    const item = add[index];
    sheet.getRange(`A${row}:O${row}`).copyFrom(sheet.getRange(`A${templateRow}:O${templateRow}`), "all");
    sheet.getRange(`A${row}:G${row}`).values = [[item.broker, item.date.value, item.ticker, item.name || null, item.action, item.units, item.price]];
    sheet.getRange(`L${row}:N${row}`).values = [[item.currency, item.fee, item.basis]];
    writeTransactionFormulas(sheet, row, item.action, item.dividendAmount);
  }
  rewriteExistingTransactionFormulas(sheet);
  refreshHoldingsSummary(sheet);
  formatStockRows(sheet, 21, Math.max(lastRow, requiredLastRow) - 20);
  formatStockReferenceBlocks(sheet);
  const currentPriceRows = readPriceRows(sheet);
  const missingCurrentPrices = holdingsRows(sheet)
    .filter(({ ticker }) => clean(currentPriceRows.get(ticker)?.price) === "")
    .map(({ ticker }) => ticker);
  const missingCostBasisRows = add.filter((item) => item.action === "Sell" && item.basis === null).map((item) => ({ ticker: item.ticker, date: item.date.text }));
  return {
    source: path.resolve(source.file), format: source.format, brokerDefault: source.defaultBroker, currencyDefault: source.defaultCurrency,
    dateBasis: parsed.dateBasis, sourceRows: parsed.sourceRows, parsedTransactions: parsed.transactions.length,
    imported: add.length, skippedDuplicates: parsed.transactions.length - add.length, ignored: parsed.ignored,
    dividendBackfills, newSecurities, missingCurrentPrices, missingCostBasisRows,
    firstRow: add.length ? firstRow : null, lastRow: add.length ? requiredLastRow : null,
  };
}

function ensureStockRows(sheet, transactions, metadata, { allowNewStockRows }) {
  const candidates = new Map();
  for (const item of transactions) {
    if (item.action === "Dividend" || metadata.has(item.ticker)) continue;
    if (!candidates.has(item.ticker)) candidates.set(item.ticker, item);
  }
  if (candidates.size && !allowNewStockRows) fail(`Ticker(s) missing from the Stock holdings summary: ${[...candidates.keys()].join(", ")}. Run import_stock.mjs with --stock-only first, then extend Asset and Summary before a full close.`);
  const additions = [];
  for (const [ticker, item] of candidates) {
    const securityType = inferSecurityType(ticker, item.name);
    if (!securityType) fail(`Cannot determine whether new ticker ${ticker} is a Stock or ETF. Add explicit metadata before importing it.`);
    const bucket = securityType === "ETF" ? "Growth" : "High Growth";
    const priceRows = readPriceRows(sheet);
    if (!priceRows.has(ticker)) {
      const targetRow = priceRows.size ? Math.max(...[...priceRows.values()].map((entry) => entry.row)) + 1 : PRICE_FIRST_ROW;
      if (targetRow > PRICE_LAST_ROW) fail(`Current-price input has no capacity for new ticker ${ticker}.`);
      sheet.getRange(`P${targetRow}:S${targetRow}`).copyFrom(sheet.getRange(`P${Math.max(PRICE_FIRST_ROW, targetRow - 1)}:S${Math.max(PRICE_FIRST_ROW, targetRow - 1)}`), "all");
      sheet.getRange(`P${targetRow}:S${targetRow}`).values = [[ticker, "", item.currency, ""]];
    }
    const existingHoldings = holdingsRows(sheet);
    const targetRow = Math.max(HOLDINGS_HEADER_ROW, ...existingHoldings.map((entry) => entry.row)) + 1;
    if (targetRow > HOLDINGS_LAST_ROW) fail(`Holdings summary has no capacity for new ticker ${ticker}.`);
    sheet.getRange(`P${targetRow}:Z${targetRow}`).copyFrom(sheet.getRange(`P${Math.max(HOLDINGS_FIRST_ROW, targetRow - 1)}:Z${Math.max(HOLDINGS_FIRST_ROW, targetRow - 1)}`), "all");
    sheet.getRange(`P${targetRow}`).values = [[ticker]];
    sheet.getRange(`Q${targetRow}`).values = [[item.name || ticker]];
    sheet.getRange(`Z${targetRow}`).values = [[bucket]];
    writeHoldingFormulas(sheet, targetRow);
    metadata.set(ticker, { name: item.name || ticker, row: targetRow, bucket, securityType });
    additions.push({ ticker, name: item.name || ticker, securityType, bucket, holdingsRow: targetRow, currentPrice: null });
  }
  return additions;
}

function refreshHoldingsSummary(sheet, bootstrap = null) {
  const entries = holdingsRows(sheet);
  if (!entries.length) return;
  const firstRow = entries[0].row;
  const lastRow = entries.at(-1).row;
  const tickers = sheet.getRange(`P${firstRow}:P${lastRow}`).values;
  sheet.getRange(`R${firstRow}:Y${lastRow}`).formulas = tickers.map((tickerRow, index) => {
    const row = firstRow + index;
    if (!normalizeTicker(tickerRow[0])) return Array(8).fill("");
    const ticker = normalizeTicker(tickerRow[0]);
    const opening = bootstrap?.openingHoldings.find((item) => item.ticker === ticker) || null;
    return holdingFormulas(row, opening, bootstrap?.openingDateText || null);
  });
}

function writeHoldingFormulas(sheet, row) {
  sheet.getRange(`R${row}:Y${row}`).formulas = [holdingFormulas(row)];
}

function holdingFormulas(row, opening = null, openingDate = null) {
  const dateCriteria = openingDate ? `,$B$21:$B$212,\">=\"&DATE(${openingDate.slice(0, 4)},${Number(openingDate.slice(5, 7))},${Number(openingDate.slice(8, 10))})` : "";
  const openingUnits = opening?.units || 0;
  const openingCost = opening?.totalCost || 0;
  const openingRealized = opening?.realized || 0;
  return [
    `=${openingUnits}+SUMIFS($F$21:$F$212,$C$21:$C$212,P${row},$E$21:$E$212,"Buy"${dateCriteria})-SUMIFS($F$21:$F$212,$C$21:$C$212,P${row},$E$21:$E$212,"Sell"${dateCriteria})`,
    `=${openingCost}+SUMIFS($H$21:$H$212,$C$21:$C$212,P${row},$E$21:$E$212,"Buy"${dateCriteria})-SUMIFS($N$21:$N$212,$C$21:$C$212,P${row},$E$21:$E$212,"Sell"${dateCriteria})`,
    `=IF(R${row}=0,0,S${row}/R${row})`,
    `=IFERROR(VLOOKUP(P${row},$P$${PRICE_FIRST_ROW}:$Q$${PRICE_LAST_ROW},2,FALSE()),"")`,
    `=IF(OR(R${row}="",U${row}=""),"",R${row}*U${row})`,
    `=IF(OR(V${row}="",S${row}=""),"",V${row}-S${row})`,
    `=${openingRealized}+SUMIFS($O$21:$O$212,$C$21:$C$212,P${row},$E$21:$E$212,"Sell"${dateCriteria})`,
    `=IF(OR(W${row}="",X${row}=""),"",W${row}+X${row})`,
  ];
}

function rewriteExistingTransactionFormulas(sheet) {
  const rows = sheet.getRange("A21:O212").values;
  let lastIndex = -1;
  for (let index = 0; index < rows.length; index += 1) if (clean(rows[index][0])) lastIndex = index;
  if (lastIndex < 0) return;
  const populated = rows.slice(0, lastIndex + 1);
  sheet.getRange(`H21:K${lastIndex + 21}`).formulas = populated.map((_, index) => transactionFormulas(index + 21));
  sheet.getRange(`O21:O${lastIndex + 21}`).formulas = populated.map((row, index) => [clean(row[4]) === "Dividend" ? "" : realizedFormula(index + 21)]);
  populated.forEach((row, index) => {
    if (clean(row[4]) === "Dividend") sheet.getRange(`O${index + 21}`).values = [[Number(row[14]) || 0]];
  });
}

function backfillClosedPositionCostBasis(sheet) {
  const rows = sheet.getRange("A21:O212").values;
  const positions = new Map();
  const inferred = [];
  for (let index = 0; index < rows.length; index += 1) {
    const row = rows[index];
    const broker = clean(row[0]);
    const ticker = normalizeTicker(row[2]);
    const action = clean(row[4]);
    if (!broker || !ticker || !["Buy", "Sell"].includes(action)) continue;
    const key = `${broker}|${ticker}`;
    const position = positions.get(key) || { units: 0, cost: 0, reliable: true };
    const units = Math.abs(Number(row[5]) || 0);
    const fee = Math.abs(Number(row[12]) || 0);
    const formulaTradeValue = units * (Number(row[6]) || 0);
    const tradeValue = Math.abs(Number(row[7]) || formulaTradeValue);
    if (action === "Buy") {
      position.units += units;
      position.cost += tradeValue + fee;
      positions.set(key, position);
      continue;
    }

    const basisPresent = clean(row[13]) !== "" && Number.isFinite(Number(row[13]));
    let basis = basisPresent ? Math.abs(Number(row[13])) : null;
    const closesPosition = position.units > 0 && Math.abs(position.units - units) <= Math.max(0.000001, position.units * 0.000001);
    if (basis === null && position.reliable && closesPosition) {
      basis = position.cost;
      sheet.getRange(`N${index + 21}`).values = [[basis]];
      inferred.push({ row: index + 21, broker, ticker, units, costBasis: basis, method: "complete-position cost roll-forward" });
    } else if (basis === null) {
      position.reliable = false;
    }
    position.units = Math.max(0, position.units - units);
    if (basis !== null) position.cost = Math.max(0, position.cost - basis);
    if (position.units <= 0.000001) {
      position.units = 0;
      position.cost = 0;
      position.reliable = true;
    }
    positions.set(key, position);
  }
  return inferred;
}

function writeTransactionFormulas(sheet, row, action, dividendAmount) {
  sheet.getRange(`H${row}:K${row}`).formulas = [transactionFormulas(row)];
  if (action === "Dividend") sheet.getRange(`O${row}`).values = [[dividendAmount || 0]];
  else sheet.getRange(`O${row}`).formulas = [[realizedFormula(row)]];
}

function transactionFormulas(row) {
  return [
    `=IF(OR(F${row}="",G${row}=""),"",F${row}*G${row})`,
    `=IFERROR(VLOOKUP(C${row},$P$${PRICE_FIRST_ROW}:$Q$${PRICE_LAST_ROW},2,FALSE()),"")`,
    `=IF(OR(F${row}="",I${row}=""),"",F${row}*I${row})`,
    `=IF(OR(J${row}="",H${row}=""),"",J${row}-H${row})`,
  ];
}

function realizedFormula(row) { return `=IF(E${row}="Sell",IF(N${row}="","",H${row}-M${row}-N${row}),0)`; }

function formatStockRows(sheet, firstRow, count) {
  if (count <= 0) return;
  const lastRow = firstRow + count - 1;
  const all = sheet.getRange(`A${firstRow}:O${lastRow}`);
  all.format.font.name = "Calibri";
  all.format.font.size = 11;
  all.format.wrapText = false;
  all.format.horizontalAlignment = "center";
  all.format.verticalAlignment = "center";
  for (const column of ["A", "B", "C", "D", "E", "F", "G", "L", "M", "N"]) {
    const range = sheet.getRange(`${column}${firstRow}:${column}${lastRow}`);
    range.format.font.color = "#008000";
    range.format.borders = { bottom: { style: "thin", color: "#D9D9D9" } };
  }
  for (const column of ["H", "I", "J", "K", "O"]) {
    const range = sheet.getRange(`${column}${firstRow}:${column}${lastRow}`);
    range.format.font.color = "#000000";
    range.format.borders = { preset: "none" };
  }
  sheet.getRange(`B${firstRow}:B${lastRow}`).format.numberFormat = "yyyy-mm-dd";
  sheet.getRange(`F${firstRow}:F${lastRow}`).format.numberFormat = "0.000000;[Red](0.000000);0.000000";
  for (const column of ["G", "H", "I", "J", "K", "M", "N", "O"]) sheet.getRange(`${column}${firstRow}:${column}${lastRow}`).format.numberFormat = "#,##0.00;[Red](#,##0.00);0.00";
}

function formatStockReferenceBlocks(sheet) {
  for (const address of [`P${PRICE_FIRST_ROW}:S${PRICE_LAST_ROW}`, `P${HOLDINGS_FIRST_ROW}:Z${HOLDINGS_LAST_ROW}`]) {
    const range = sheet.getRange(address);
    range.format.font.name = "Calibri";
    range.format.font.size = 11;
    range.format.wrapText = false;
    range.format.horizontalAlignment = "center";
    range.format.verticalAlignment = "center";
  }
  sheet.getRange(`Q${PRICE_FIRST_ROW}:Q${PRICE_LAST_ROW}`).format.numberFormat = "#,##0.00;[Red](#,##0.00);0.00";
  sheet.getRange(`S${PRICE_FIRST_ROW}:S${PRICE_LAST_ROW}`).format.numberFormat = "yyyy-mm-dd";
  sheet.getRange(`R${HOLDINGS_FIRST_ROW}:R${HOLDINGS_LAST_ROW}`).format.numberFormat = "0.000000;[Red](0.000000);0.000000";
  for (const column of ["S", "T", "U", "V", "W", "X", "Y"]) sheet.getRange(`${column}${HOLDINGS_FIRST_ROW}:${column}${HOLDINGS_LAST_ROW}`).format.numberFormat = "#,##0.00;[Red](#,##0.00);0.00";
  sheet.getRange("D20:D212").format.columnWidth = 24;
  sheet.getRange(`Q${HOLDINGS_HEADER_ROW}:Q${HOLDINGS_LAST_ROW}`).format.columnWidth = 24;
}

function normalizeSecurityNames(sheet) {
  const transactions = sheet.getRange("C21:D212").values;
  sheet.getRange("D21:D212").values = transactions.map(([ticker, name]) => {
    const normalizedTicker = normalizeTicker(ticker);
    return [normalizedTicker ? canonicalSecurityName(normalizedTicker, name) : ""];
  });
  const holdings = sheet.getRange(`P${HOLDINGS_FIRST_ROW}:Q${HOLDINGS_LAST_ROW}`).values;
  sheet.getRange(`Q${HOLDINGS_FIRST_ROW}:Q${HOLDINGS_LAST_ROW}`).values = holdings.map(([ticker, name]) => {
    const normalizedTicker = normalizeTicker(ticker);
    return [normalizedTicker ? canonicalSecurityName(normalizedTicker, name) : ""];
  });
}

function sortTransactionLedger(sheet) {
  const values = sheet.getRange("A21:O212").values;
  const populated = values
    .map((row, index) => ({ row, index, date: excelDateText(row[1]) }))
    .filter((item) => clean(item.row[0]))
    .sort((left, right) => left.date.localeCompare(right.date) || left.index - right.index);
  if (!populated.length) return { rowCount: 0, firstDate: null, lastDate: null };

  const lastRow = 20 + populated.length;
  sheet.getRange("A21:O212").clear({ applyTo: "contents" });
  sheet.getRange(`A21:G${lastRow}`).values = populated.map(({ row }) => row.slice(0, 7).map(blankValue));
  sheet.getRange(`L21:N${lastRow}`).values = populated.map(({ row }) => row.slice(11, 14).map(blankValue));
  for (let index = 0; index < populated.length; index += 1) {
    const source = populated[index].row;
    if (clean(source[4]) === "Dividend") sheet.getRange(`O${index + 21}`).values = [[Number(source[14]) || 0]];
  }
  rewriteExistingTransactionFormulas(sheet);
  return { rowCount: populated.length, firstDate: populated[0].date, lastDate: populated.at(-1).date };
}

async function fetchAndApplyCurrentPrices(sheet, inputDir, priceAsOf) {
  const python = process.env.WEALTH_PYTHON;
  if (!python) fail("WEALTH_PYTHON is required to fetch current stock prices.");
  const tickers = [...readPriceRows(sheet).keys()];
  if (!tickers.length) fail("Stock current-price input has no tickers.");

  const temporaryDirectory = await fs.mkdtemp(path.join(os.tmpdir(), "wealth-stock-prices-"));
  const priceCsv = path.join(temporaryDirectory, "prices.csv");
  const script = path.join(path.dirname(fileURLToPath(import.meta.url)), "fetch_stock_prices.py");
  const args = [script, "--output", priceCsv, "--tickers", tickers.join(",")];
  if (priceAsOf) args.push("--as-of", priceAsOf);
  else {
    const cashflowDirectory = path.join(path.resolve(inputDir), "CashFlow");
    const cashflowFiles = await discoverCsvFiles(cashflowDirectory);
    if (!cashflowFiles.length) fail("No CashFlow CSV is available to determine the current-price date. Supply --price-as-of YYYY-MM-DD.");
    for (const file of cashflowFiles) args.push("--cashflow-csv", file);
  }

  try {
    const { stdout } = await execFileAsync(python, args, { maxBuffer: 1024 * 1024 * 10 });
    const fetchReport = JSON.parse(stdout.trim().split(/\r?\n/).at(-1));
    const rows = await readCsvRecords(priceCsv);
    const latest = new Map();
    for (const row of rows.records) {
      const ticker = normalizeTicker(row.Ticker);
      const date = parseDate(row.Date);
      if (!ticker || !date) continue;
      const current = latest.get(ticker);
      if (!current || date.text > current.date.text) latest.set(ticker, { price: parseNumber(row["Current Price"], "current price"), currency: clean(row.Currency) || "USD", date });
    }
    const priceRows = readPriceRows(sheet);
    const missing = [];
    for (const ticker of tickers) {
      const fetched = latest.get(ticker);
      if (!fetched) { missing.push(ticker); continue; }
      const targetRow = priceRows.get(ticker).row;
      sheet.getRange(`Q${targetRow}:S${targetRow}`).values = [[fetched.price, fetched.currency, fetched.date.value]];
    }
    if (missing.length) fail(`Price fetch did not return current prices for: ${missing.join(", ")}`);
    return { pricingDate: fetchReport.pricing_date, tickers, source: "fetch_stock_prices.py" };
  } finally {
    await fs.rm(temporaryDirectory, { recursive: true, force: true });
  }
}

async function discoverCsvFiles(directory) {
  if (!await exists(directory)) return [];
  const entries = await fs.readdir(directory, { withFileTypes: true });
  return entries.filter((entry) => entry.isFile() && path.extname(entry.name).toLowerCase() === ".csv").map((entry) => path.join(directory, entry.name)).sort();
}

function readTickerMetadata(sheet) {
  const result = new Map();
  const holdings = sheet.getRange(`P${HOLDINGS_FIRST_ROW}:Z${HOLDINGS_LAST_ROW}`).values;
  for (let index = 0; index < holdings.length; index += 1) {
    const ticker = normalizeTicker(holdings[index][0]);
    if (ticker) result.set(ticker, { row: index + HOLDINGS_FIRST_ROW, name: clean(holdings[index][1]), bucket: clean(holdings[index][10]) });
  }
  const ledger = sheet.getRange("C21:D212").values;
  for (const row of ledger) {
    const ticker = normalizeTicker(row[0]);
    if (ticker && !result.has(ticker)) result.set(ticker, { row: null, name: clean(row[1]), bucket: "" });
  }
  return result;
}

function holdingsRows(sheet) {
  return sheet.getRange(`P${HOLDINGS_FIRST_ROW}:P${HOLDINGS_LAST_ROW}`).values.map((row, index) => ({ row: index + HOLDINGS_FIRST_ROW, ticker: normalizeTicker(row[0]) })).filter((item) => item.ticker);
}

function readPriceRows(sheet) {
  const result = new Map();
  const values = sheet.getRange(`P${PRICE_FIRST_ROW}:S${PRICE_LAST_ROW}`).values;
  for (let index = 0; index < values.length; index += 1) {
    const ticker = normalizeTicker(values[index][0]);
    if (ticker) result.set(ticker, { row: index + PRICE_FIRST_ROW, price: values[index][1], currency: clean(values[index][2]) });
  }
  return result;
}

function sameTransaction(existing, incoming) {
  return canonicalBroker(existing[0]) === incoming.broker
    && excelDateText(existing[1]) === incoming.date.text
    && normalizeTicker(existing[2]) === incoming.ticker
    && clean(existing[4]) === incoming.action
    && clean(existing[11]).toUpperCase() === incoming.currency
    && Math.abs(Number(existing[5] || 0) - incoming.units) <= 0.0001
    && Math.abs(Number(existing[6] || 0) - incoming.price) <= 0.01;
}

function findCoveredTransactionGroups(existingRows, incomingTransactions) {
  const existingGroups = new Map();
  for (const row of existingRows) {
    if (!clean(row[0])) continue;
    const item = {
      broker: canonicalBroker(row[0]), date: { text: excelDateText(row[1]) }, ticker: normalizeTicker(row[2]),
      action: clean(row[4]), currency: clean(row[11]).toUpperCase(),
    };
    const key = transactionGroupKey(item);
    const group = existingGroups.get(key) || { units: 0, tradeValue: 0 };
    group.units += Number(row[5] || 0);
    group.tradeValue += Number(row[7] || 0);
    existingGroups.set(key, group);
  }

  const incomingGroups = new Map();
  for (const item of incomingTransactions) {
    if (item.action === "Dividend") continue;
    const key = transactionGroupKey(item);
    const group = incomingGroups.get(key) || { units: 0, tradeValue: 0 };
    group.units += item.units;
    group.tradeValue += item.units * item.price;
    incomingGroups.set(key, group);
  }

  const covered = new Set();
  for (const [key, incoming] of incomingGroups) {
    const existing = existingGroups.get(key);
    if (!existing) continue;
    if (Math.abs(existing.units - incoming.units) <= 0.0002
      && Math.abs(existing.tradeValue - incoming.tradeValue) <= 0.10) covered.add(key);
  }
  return covered;
}

function transactionGroupKey(item) {
  return JSON.stringify([item.broker, item.date.text, item.ticker, item.action, item.currency]);
}

function aggregateTransactions(transactions) {
  const groups = new Map();
  for (const item of transactions) {
    const key = JSON.stringify([item.broker, item.date.text, item.ticker, item.action, round(item.price, 6), item.currency]);
    if (!groups.has(key)) groups.set(key, { ...item });
    else {
      const existing = groups.get(key);
      existing.units += item.units;
      existing.fee = existing.fee === null && item.fee === null ? null : (existing.fee || 0) + (item.fee || 0);
      existing.basis = existing.basis === null || item.basis === null ? null : existing.basis + item.basis;
      existing.dividendAmount = existing.dividendAmount === null && item.dividendAmount === null ? null : (existing.dividendAmount || 0) + (item.dividendAmount || 0);
      if (!existing.name && item.name) existing.name = item.name;
    }
  }
  return [...groups.values()].map((item) => ({ ...item, units: round(item.units, 9), fee: item.fee === null ? null : round(item.fee, 9), basis: item.basis === null ? null : round(item.basis, 9), dividendAmount: item.dividendAmount === null ? null : round(item.dividendAmount, 9) }));
}

function groupPdfLines(items) {
  const lines = [];
  for (const item of items) {
    const y = Math.round(item.transform[5] * 2) / 2;
    let line = lines.find((candidate) => Math.abs(candidate.y - y) < 1);
    if (!line) { line = { y, items: [] }; lines.push(line); }
    line.items.push({ x: item.transform[4], text: item.str });
  }
  return lines.sort((a, b) => b.y - a.y).map((line) => ({ y: line.y, tokens: line.items.sort((a, b) => a.x - b.x).map((item) => item.text) }));
}

function previousWeekday(date) {
  const value = new Date(date.value.getTime() - 86400000);
  while ([0, 6].includes(value.getUTCDay())) value.setUTCDate(value.getUTCDate() - 1);
  return parseDate(value.toISOString().slice(0, 10));
}

function inferSecurityType(ticker, name) {
  if (KNOWN_ETFS.has(ticker) || /\bETF\b|EXCHANGE[- ]TRADED|INDEX FUND/i.test(name)) return "ETF";
  if (/^[A-Z][A-Z0-9.\/-]{0,9}$/.test(ticker) && name) return "Stock";
  return null;
}

export function canonicalSecurityName(ticker, fallback) {
  const canonicalTicker = normalizeTicker(ticker);
  const mapped = SECURITY_NAMES.get(canonicalTicker);
  if (mapped) return mapped;
  const sourceName = clean(fallback).split(/\r?\n/)[0];
  if (!sourceName || /^(CASH\s+DIV|DIVIDEND\b)/i.test(sourceName)) return canonicalTicker;
  return sourceName;
}

function canonicalAction(value) {
  const normalized = clean(value).toUpperCase();
  if (["BUY", "BOUGHT", "YOU BOUGHT", "PURCHASE", "REINVEST SHARES", "买入"].includes(normalized)) return "Buy";
  if (["SELL", "SOLD", "YOU SOLD", "卖出"].includes(normalized)) return "Sell";
  if (["DIVIDEND", "DIVIDEND|", "DIV", "CDIV", "DIVIDEND RECEIVED", "CASH DIVIDEND", "QUALIFIED DIVIDEND", "股息"].includes(normalized)) return "Dividend";
  return null;
}

function isCashActivity(action, description) {
  const text = `${clean(action)} ${clean(description)}`.toUpperCase();
  return /\b(ACH|DCF|CASH|TRANSFER|DEP|DEPOSIT|WITHDRAWAL|WIRE|JOURNAL|MONEY MARKET)\b/.test(text);
}

function canonicalBroker(value) {
  const normalized = clean(value).toLowerCase();
  if (["charles", "charles schwab", "schwab"].includes(normalized)) return "Charles";
  if (normalized === "interactive brokers") return "IBKR";
  if (normalized === "fidelity") return "Fidelity";
  if (normalized === "robinhood") return "Robinhood";
  if (normalized === "ibkr") return "IBKR";
  return clean(value);
}

function inferSource(file) {
  const resolved = path.resolve(typeof file === "string" ? file : file.file);
  const folder = path.basename(path.dirname(resolved));
  const defaults = BROKER_FOLDERS[folder];
  if (!defaults) fail(`Cannot infer broker from source path: ${resolved}`);
  const extension = path.extname(resolved).toLowerCase();
  const format = extension === ".pdf" && folder === "Fidelity" ? "fidelity-pdf" : extension === ".csv" ? "csv" : null;
  if (!format) fail(`Unsupported brokerage source: ${resolved}`);
  return { file: resolved, format, defaultBroker: defaults.broker, defaultCurrency: defaults.currency };
}

function looksLikeStockCsv(headers, defaults) {
  const has = (...names) => names.some((name) => headers.includes(name));
  return has("Date", "Trade Date", "Activity Date", "Date/Time", "日期")
    && has("Ticker", "Symbol", "Instrument", "代码", "股票代码")
    && has("Action", "Side", "Type", "Trans Code", "Buy/Sell", "操作", "买卖")
    && (has("Currency", "币种", "货币") || defaults.currency);
}

async function readCsvRecords(csvPath) {
  const text = (await fs.readFile(csvPath, "utf8")).replace(/^\uFEFF/, "");
  const workbook = await Workbook.fromCSV(text, { sheetName: "CSV" });
  const values = workbook.worksheets.getItem("CSV").getUsedRange(true).values;
  if (!values?.length) fail(`CSV is empty: ${csvPath}`);
  const headers = values[0].map(clean);
  const records = values.slice(1).filter((row) => row.some((value) => clean(value) !== "")).map((row, index) => {
    const record = { __row: index + 2 };
    headers.forEach((header, column) => { record[header] = row[column] ?? ""; });
    return record;
  });
  return { headers, records };
}

async function readCsvHeaders(csvPath) {
  const text = (await fs.readFile(csvPath, "utf8")).replace(/^\uFEFF/, "");
  const firstLine = text.split(/\r?\n/).find((line) => line.trim() !== "");
  if (!firstLine) return [];
  const values = [];
  let value = "";
  let quoted = false;
  for (let index = 0; index < firstLine.length; index += 1) {
    const character = firstLine[index];
    if (character === '"' && quoted && firstLine[index + 1] === '"') { value += '"'; index += 1; }
    else if (character === '"') quoted = !quoted;
    else if (character === "," && !quoted) { values.push(clean(value)); value = ""; }
    else value += character;
  }
  values.push(clean(value));
  return values;
}

function prepareStockOnlyPreview(workbook) {
  const stock = workbook.worksheets.getItem("Stock");
  const baseCurrency = clean(stock.getRange("B9").values[0][0]) || "USD";
  const exchangeRate = Number(stock.getRange("B10").values[0][0]) || 1;
  stock.getRange("B9").values = [[baseCurrency]];
  stock.getRange("B10").values = [[exchangeRate]];
  for (const sheetName of listSheetNames(workbook).reverse()) if (sheetName !== "Stock") workbook.worksheets.getItem(sheetName).delete();
}

function summarizeReports(reports) {
  return reports.reduce((total, report) => {
    total.sourceFiles += 1;
    total.sourceRows += report.sourceRows;
    total.parsedTransactions += report.parsedTransactions;
    total.imported += report.imported;
    total.skippedDuplicates += report.skippedDuplicates;
    total.newSecurities.push(...report.newSecurities);
    total.missingCurrentPrices.push(...report.missingCurrentPrices);
    total.missingCostBasisRows.push(...report.missingCostBasisRows);
    return total;
  }, { sourceFiles: 0, sourceRows: 0, parsedTransactions: 0, imported: 0, skippedDuplicates: 0, newSecurities: [], missingCurrentPrices: [], missingCostBasisRows: [] });
}

function summarizeStatementCache(reports, cacheDir, enabled) {
  const statuses = { hit: 0, miss: 0, refreshed: 0, repaired: 0, disabled: 0 };
  for (const report of reports) {
    const status = report.cache?.status || "disabled";
    statuses[status] = (statuses[status] || 0) + 1;
  }
  return { enabled, directory: enabled ? cacheDir : null, ...statuses };
}

function assertHeaders(sheet, range, expected) {
  const actual = sheet.getRange(range).values[0].map(clean);
  if (JSON.stringify(actual) !== JSON.stringify(expected)) fail(`Unexpected headers at ${sheet.name}!${range}: ${actual.join(" | ")}`);
}

function parseDate(value) {
  const datePart = clean(value).split(/[T\s]/)[0];
  const iso = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(datePart);
  const us = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(datePart);
  if (!iso && !us) return null;
  const [year, month, day] = iso ? [iso[1], iso[2], iso[3]] : [us[3], us[1], us[2]];
  const text = `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
  return { text, value: new Date(`${text}T00:00:00.000Z`) };
}

function parseNumber(value, label) {
  const original = clean(value);
  if (!original || original === "-") fail(`Invalid ${label}: ${value}`);
  const parenthesized = /^\(.*\)$/.test(original);
  const normalized = original.replace(/[()$,]/g, "").trim();
  const number = Number(normalized) * (parenthesized ? -1 : 1);
  if (!Number.isFinite(number)) fail(`Invalid ${label}: ${value}`);
  return number;
}

function optionalNumber(value, fallback) { return clean(value) && clean(value) !== "-" ? parseNumber(value, "number") : fallback; }
function blankValue(value) { return value === null || value === undefined ? "" : value; }
function normalizeTicker(value) { const ticker = clean(value).toUpperCase(); return ticker === "ECHO" ? "SATS" : ticker; }
function normalizeSecurityName(value) { return clean(value).toUpperCase().replace(/\s+/g, " ").replace(/\s+USD$/, ""); }
function excelDateText(value) { if (value instanceof Date) return value.toISOString().slice(0, 10); if (typeof value === "number") return new Date(Date.UTC(1899, 11, 30) + Math.round(value) * 86400000).toISOString().slice(0, 10); return clean(value).slice(0, 10); }
function round(value, digits) { const factor = 10 ** digits; return Math.round((Number(value) + Number.EPSILON) * factor) / factor; }
function clean(value) { return value === null || value === undefined ? "" : String(value).trim(); }
function listSheetNames(workbook) { const names = []; for (let index = 0; ; index += 1) { try { const sheet = workbook.worksheets.getItemAt(index); if (!sheet) break; names.push(sheet.name); } catch { break; } } return names; }
async function exists(file) { try { await fs.access(file); return true; } catch { return false; } }
function fail(message) { throw new Error(message); }

function parseCliArgs(argv) {
  const args = { stockFiles: [] };
  for (let index = 0; index < argv.length; index += 1) {
    const key = argv[index];
    if (key === "--dry-run") args.dryRun = true;
    else if (key === "--allow-existing-output") args.allowExistingOutput = true;
    else if (key === "--stock-only") args.stockOnly = true;
    else if (key === "--allow-new-stock-rows") args.allowNewStockRows = true;
    else if (key === "--skip-price-fetch") args.skipPriceFetch = true;
    else if (key === "--no-statement-cache") args.noStatementCache = true;
    else if (key === "--refresh-statement-cache") args.refreshStatementCache = true;
    else if (key === "--stock-file") args.stockFiles.push(argv[++index]);
    else if (key.startsWith("--")) args[key.slice(2).replace(/-([a-z])/g, (_, character) => character.toUpperCase())] = argv[++index];
    else fail(`Unexpected argument: ${key}`);
  }
  return args;
}

async function main() {
  const args = parseCliArgs(process.argv.slice(2));
  const inputDir = path.resolve(args.inputDir || process.cwd());
  const workbookPath = path.resolve(args.workbook || path.join(inputDir, "Output", "One_Piece.xlsx"));
  if (!args.output) fail("Usage: import_stock.mjs [--input-dir DIR] [--input-file FILE] [--workbook FILE] --output FILE [--stock-file FILE ...] [--stock-only] [--allow-new-stock-rows] [--price-as-of YYYY-MM-DD] [--skip-price-fetch] [--statement-cache-dir DIR] [--no-statement-cache] [--refresh-statement-cache] [--dry-run] [--allow-existing-output]");
  const report = await runStockImport({
    workbookPath, outputPath: args.output, inputDir,
    inputFile: args.inputFile || path.join(inputDir, "monthly-close-input.md"),
    stockSources: args.stockFiles, dryRun: Boolean(args.dryRun),
    allowExistingOutput: Boolean(args.allowExistingOutput), stockOnly: Boolean(args.stockOnly),
    allowNewStockRows: Boolean(args.allowNewStockRows),
    fetchPrices: !args.skipPriceFetch, priceAsOf: args.priceAsOf || null,
    cacheDir: args.statementCacheDir || null,
    useCache: !args.noStatementCache,
    refreshCache: Boolean(args.refreshStatementCache),
  });
  console.log(JSON.stringify(report, null, 2));
}

const directRun = process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href;
if (directRun) main().catch((error) => { console.error(error instanceof Error ? error.message : String(error)); process.exitCode = 1; });

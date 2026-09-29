import fs from "node:fs/promises";
import path from "node:path";
import { createHash } from "node:crypto";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";

const execFileAsync = promisify(execFile);
const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const CASHFLOW_CACHE_VERSION = 1;
const CASHFLOW_PARSER_VERSION = "2026-09-29.1";

const FIXED_SOURCE_FOLDERS = new Map([
  ["Alipay", { parser: "alipay", account: "支付宝", currency: "CNY" }],
  ["WeChat", { parser: "wechat", account: "WeChat", currency: "CNY" }],
  ["CITIC_CNY", { parser: "citic", account: "中信银行", currency: "CNY" }],
  ["CITIC_USD", { parser: "citic", account: "中信银行", currency: "USD" }],
]);

const SUPPORTED_EXTENSIONS = new Set([".csv", ".xlsx", ".xls"]);

export async function discoverAccountCashflowSources(cashflowDir) {
  const directory = path.resolve(cashflowDir);
  const recognized = [];
  const ignored = [];
  const rootEntries = await fs.readdir(directory, { withFileTypes: true });

  for (const rootEntry of rootEntries.filter((entry) => entry.isDirectory()).sort((a, b) => a.name.localeCompare(b.name))) {
    const folder = rootEntry.name;
    const config = sourceFolderConfig(folder);
    if (!config) continue;
    const folderPath = path.join(directory, folder);
    const entries = await fs.readdir(folderPath, { withFileTypes: true });
    for (const entry of entries.sort((a, b) => a.name.localeCompare(b.name))) {
      if (!entry.isFile() || entry.name.startsWith("~$")) continue;
      const extension = path.extname(entry.name).toLowerCase();
      const source = { file: path.join(folderPath, entry.name), folder, ...config };
      if (SUPPORTED_EXTENSIONS.has(extension)) recognized.push(source);
      else ignored.push({ file: source.file, reason: `unsupported extension: ${extension || "none"}` });
    }
  }
  return { directory, recognized, ignored };
}

function sourceFolderConfig(folder) {
  const checking = /^Chase_Checking_(\d{4})$/.exec(folder);
  if (checking) return { parser: "chase-checking", account: `Chase Checking ${checking[1]}`, currency: "USD" };
  const credit = /^Chase_Credit_(\d{4})$/.exec(folder);
  if (credit) return { parser: "chase-credit", account: `Chase Credit ${credit[1]}`, currency: "USD" };
  return FIXED_SOURCE_FOLDERS.get(folder) || null;
}

export async function readCashflowRules(cashflowDir) {
  const file = path.join(path.resolve(cashflowDir), "cashflow-rules.csv");
  try {
    const text = await fs.readFile(file, "utf8");
    const rows = parseCsv(text.replace(/^\uFEFF/, ""));
    if (!rows.length) return { file, rules: [] };
    const headers = rows[0].map(clean);
    const rules = rows.slice(1).filter(hasContent).map((row, index) => {
      const record = Object.fromEntries(headers.map((header, column) => [header, clean(row[column])]));
      return {
        ...record,
        __row: index + 2,
        amount: record.amount === "" ? null : parseAmount(record.amount),
        fx_rate: record.fx_rate === "" ? null : Number(record.fx_rate),
      };
    });
    return { file, rules };
  } catch (error) {
    if (error?.code === "ENOENT") return { file, rules: [] };
    throw error;
  }
}

export async function parseAccountCashflowSources(sources, {
  FileBlob,
  SpreadsheetFile,
  throughDate,
  rules = [],
  cacheDir = null,
  useCache = true,
  refreshCache = false,
} = {}) {
  const cutoff = parseDateText(throughDate);
  if (!cutoff) throw new Error(`Invalid CashFlow cutoff date: ${throughDate}`);
  const items = { income: [], expense: [], transfer: [] };
  const files = [];
  const unresolved = [];
  const usedRules = new Set();
  const parsedFiles = [];

  for (const source of sources) {
    const { transactions: rawTransactions, cache } = await readOrParseCashflowSource(source, {
      FileBlob,
      SpreadsheetFile,
      cacheDir,
      useCache,
      refreshCache,
    });
    const accepted = rawTransactions.filter((transaction) => transaction.date <= cutoff);
    const excludedAfterCutoff = rawTransactions.length - accepted.length;
    parsedFiles.push({ source, rawTransactions, accepted, excludedAfterCutoff, cache });
  }

  const reconciliation = reconcileTransactionOccurrences(parsedFiles);
  for (const parsedFile of parsedFiles) {
    const { source, rawTransactions, accepted, excludedAfterCutoff, cache } = parsedFile;
    const reconciled = reconciliation.byFile.get(source.file) || [];
    const sourceCounts = { income: 0, expense: 0, transfer: 0, ignored: 0 };

    for (const transaction of reconciled) {
      const match = rules.find((rule) => ruleMatches(rule, transaction));
      if (match) usedRules.add(match.__row);
      const result = classifyTransaction(transaction, match);
      if (!result) {
        unresolved.push({
          file: source.file,
          row: transaction.sourceRow,
          date: transaction.date,
          amount: transaction.signedAmount,
          type: transaction.type,
          description: transaction.description,
        });
        continue;
      }
      if (result.kind === "ignore") {
        sourceCounts.ignored += 1;
        continue;
      }
      const item = toCashflowItem(transaction, result);
      items[result.kind].push(item);
      sourceCounts[result.kind] += 1;
      if (result.kind === "transfer" && result.mirrorAccount) {
        items.transfer.push(toCashflowItem({ ...transaction, account: result.mirrorAccount, signedAmount: -transaction.signedAmount }, { ...result, mirrorAccount: "" }));
        sourceCounts.transfer += 1;
      }
    }

    files.push({
      source: source.file,
      sourceRows: rawTransactions.length,
      acceptedRows: accepted.length,
      reconciledRows: reconciled.length,
      skippedDuplicateOccurrences: accepted.length - reconciled.length,
      excludedAfterCutoff,
      latestDate: accepted.map((transaction) => transaction.date).sort().at(-1) || null,
      counts: sourceCounts,
      cache,
    });
  }

  if (unresolved.length) {
    throw new Error(`CashFlow has unresolved transactions. Add explicit rules to CashFlow/cashflow-rules.csv:\n${JSON.stringify(unresolved.slice(0, 30), null, 2)}`);
  }

  for (const kind of Object.keys(items)) {
    items[kind].sort((a, b) => a.date.localeCompare(b.date) || a.source.localeCompare(b.source) || a.sourceRow - b.sourceRow);
  }
  return {
    items,
    files,
    duplicateOccurrences: reconciliation.duplicateOccurrences,
    usedRuleRows: [...usedRules].sort((a, b) => a - b),
  };
}

async function readOrParseCashflowSource(source, { FileBlob, SpreadsheetFile, cacheDir, useCache, refreshCache }) {
  if (!useCache) {
    return {
      transactions: await parseSource(source, { FileBlob, SpreadsheetFile }),
      cache: { status: "disabled", file: null },
    };
  }
  if (!cacheDir) throw new Error("CashFlow cache directory is required when caching is enabled.");

  const sourceBytes = await fs.readFile(source.file);
  const sourceHash = createHash("sha256").update(sourceBytes).digest("hex");
  const cacheKey = createHash("sha256").update(JSON.stringify({
    sourceHash,
    parser: source.parser,
    account: source.account,
    currency: source.currency,
    parserVersion: CASHFLOW_PARSER_VERSION,
  })).digest("hex");
  const cacheFile = path.join(cacheDir, `${cacheKey}.json`);
  let invalidCache = false;

  if (!refreshCache && await exists(cacheFile)) {
    try {
      const cached = JSON.parse(await fs.readFile(cacheFile, "utf8"));
      return {
        transactions: hydrateCachedTransactions(cached, sourceHash, source),
        cache: { status: "hit", file: cacheFile, sourceHash },
      };
    } catch {
      invalidCache = true;
    }
  }

  const transactions = await parseSource(source, { FileBlob, SpreadsheetFile });
  const payload = {
    cacheVersion: CASHFLOW_CACHE_VERSION,
    parserVersion: CASHFLOW_PARSER_VERSION,
    sourceHash,
    parser: source.parser,
    account: source.account,
    currency: source.currency,
    transactions: transactions.map(serializeTransaction),
  };
  await fs.mkdir(cacheDir, { recursive: true });
  const temporaryFile = `${cacheFile}.${process.pid}.tmp`;
  await fs.writeFile(temporaryFile, `${JSON.stringify(payload)}\n`, "utf8");
  await fs.rename(temporaryFile, cacheFile);
  return {
    transactions,
    cache: { status: refreshCache ? "refreshed" : invalidCache ? "repaired" : "miss", file: cacheFile, sourceHash },
  };
}

function serializeTransaction(transaction) {
  const { source: _source, sourceFolder: _sourceFolder, account: _account, currency: _currency, ...portable } = transaction;
  return portable;
}

function hydrateCachedTransactions(cached, sourceHash, source) {
  if (cached?.cacheVersion !== CASHFLOW_CACHE_VERSION
    || cached?.parserVersion !== CASHFLOW_PARSER_VERSION
    || cached?.sourceHash !== sourceHash
    || cached?.parser !== source.parser
    || cached?.account !== source.account
    || cached?.currency !== source.currency
    || !Array.isArray(cached?.transactions)) throw new Error("Invalid CashFlow cache entry.");
  return cached.transactions.map((transaction) => {
    if (!parseDateText(transaction.date) || !Number.isFinite(transaction.signedAmount)) throw new Error("Invalid cached CashFlow transaction.");
    return {
      ...transaction,
      source: source.file,
      sourceFolder: source.folder,
      account: source.account,
      currency: source.currency,
    };
  });
}

function reconcileTransactionOccurrences(parsedFiles) {
  const selectedByKey = new Map();
  let acceptedOccurrences = 0;

  for (const parsedFile of parsedFiles) {
    acceptedOccurrences += parsedFile.accepted.length;
    const fileGroups = new Map();
    for (const transaction of parsedFile.accepted) {
      const key = transactionOccurrenceKey(transaction);
      if (!fileGroups.has(key)) fileGroups.set(key, []);
      fileGroups.get(key).push(transaction);
    }
    for (const [key, occurrences] of fileGroups) {
      const selected = selectedByKey.get(key);
      if (!selected || occurrences.length > selected.length) selectedByKey.set(key, occurrences);
    }
  }

  const byFile = new Map(parsedFiles.map(({ source }) => [source.file, []]));
  let reconciledOccurrences = 0;
  for (const occurrences of selectedByKey.values()) {
    reconciledOccurrences += occurrences.length;
    for (const transaction of occurrences) byFile.get(transaction.source).push(transaction);
  }
  return {
    byFile,
    duplicateOccurrences: acceptedOccurrences - reconciledOccurrences,
  };
}

function transactionOccurrenceKey(transaction) {
  return JSON.stringify([
    transaction.account,
    transaction.currency,
    transaction.date,
    transaction.signedAmount,
    transaction.type,
    transaction.description,
    transaction.rawCategory,
    transaction.direction,
  ]);
}

async function parseSource(source, runtime) {
  if (source.parser === "chase-checking") return parseChaseChecking(source);
  if (source.parser === "chase-credit") return parseChaseCredit(source);
  if (source.parser === "alipay") return parseAlipay(source);
  if (source.parser === "wechat") return parseWeChat(source, runtime);
  if (source.parser === "citic") return parseCitic(source, runtime);
  throw new Error(`Unsupported CashFlow parser: ${source.parser}`);
}

async function parseChaseChecking(source) {
  const rows = parseCsv((await fs.readFile(source.file, "utf8")).replace(/^\uFEFF/, ""));
  const records = recordsFromRows(rows, ["Posting Date", "Description", "Amount", "Type"], source.file);
  return records.map((record) => normalizedTransaction(source, record, {
    date: record["Posting Date"],
    amount: record.Amount,
    type: record.Type,
    description: record.Description,
    category: "",
  }));
}

async function parseChaseCredit(source) {
  const rows = parseCsv((await fs.readFile(source.file, "utf8")).replace(/^\uFEFF/, ""));
  const records = recordsFromRows(rows, ["Transaction Date", "Description", "Category", "Type", "Amount"], source.file);
  return records.map((record) => normalizedTransaction(source, record, {
    date: record["Transaction Date"],
    amount: record.Amount,
    type: record.Type,
    description: record.Description,
    category: record.Category,
  }));
}

async function parseAlipay(source) {
  const bytes = await fs.readFile(source.file);
  const text = new TextDecoder("gb18030").decode(bytes).replace(/^\uFEFF/, "");
  const rows = parseCsv(text);
  const headerIndex = rows.findIndex((row) => row.map(clean).includes("交易时间") && row.map(clean).includes("收/支"));
  if (headerIndex < 0) throw new Error(`Could not find Alipay transaction headers in ${source.file}`);
  const records = recordsFromRows(rows.slice(headerIndex), ["交易时间", "交易分类", "交易对方", "商品说明", "收/支", "金额"], source.file, headerIndex);
  return records.filter((record) => ["收入", "支出", "不计收支"].includes(clean(record["收/支"]))).map((record) => {
    const direction = clean(record["收/支"]);
    const amount = Math.abs(parseAmount(record["金额"]));
    return normalizedTransaction(source, record, {
      date: record["交易时间"],
      amount: direction === "支出" ? -amount : amount,
      type: record["交易分类"],
      description: [record["交易对方"], record["商品说明"]].map(clean).filter(Boolean).join(" - "),
      category: "",
      direction,
    });
  });
}

async function parseWeChat(source, { FileBlob, SpreadsheetFile }) {
  if (path.extname(source.file).toLowerCase() !== ".xlsx") throw new Error(`WeChat source must be .xlsx: ${source.file}`);
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(source.file));
  const sheet = workbook.worksheets.getItemAt(0);
  const rows = sheet.getUsedRange(true).values;
  const headerIndex = rows.findIndex((row) => row.map(clean).includes("交易时间") && row.map(clean).includes("收/支"));
  if (headerIndex < 0) throw new Error(`Could not find WeChat transaction headers in ${source.file}`);
  const records = recordsFromRows(rows.slice(headerIndex), ["交易时间", "交易类型", "交易对方", "商品", "收/支", "金额(元)"], source.file, headerIndex);
  return records.filter((record) => ["收入", "支出", "不计收支"].includes(clean(record["收/支"]))).map((record) => {
    const direction = clean(record["收/支"]);
    const amount = Math.abs(parseAmount(record["金额(元)"]));
    return normalizedTransaction(source, record, {
      date: record["交易时间"],
      amount: direction === "支出" ? -amount : amount,
      type: record["交易类型"],
      description: [record["交易对方"], record["商品"]].map(clean).filter(Boolean).join(" - "),
      category: "",
      direction,
    });
  });
}

async function parseCitic(source, { FileBlob, SpreadsheetFile }) {
  const extension = path.extname(source.file).toLowerCase();
  if (extension === ".xlsx") {
    const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(source.file));
    return parseCiticRows(source, workbook.worksheets.getItemAt(0).getUsedRange(true).values);
  }
  if (extension === ".xls") {
    const python = process.env.WEALTH_PYTHON;
    if (!python) throw new Error("WEALTH_PYTHON is required to read legacy CITIC .xls files.");
    const helper = path.join(scriptDirectory, "read_legacy_xls.py");
    let stdout;
    try {
      ({ stdout } = await execFileAsync(python, [helper, source.file], { maxBuffer: 10 * 1024 * 1024 }));
    } catch (error) {
      throw new Error(`Could not read CITIC .xls file. Install skills/update-treasury/requirements.txt into WEALTH_PYTHON. ${error.stderr || error.message}`);
    }
    return parseCiticRows(source, JSON.parse(stdout));
  }
  throw new Error(`Unsupported CITIC file type: ${source.file}`);
}

function parseCiticRows(source, rows) {
  const headerIndex = rows.findIndex((row) => row.map(clean).includes("交易日期") && row.map(clean).some((value) => /收入金额|支出金额|交易金额/.test(value)));
  if (headerIndex < 0) throw new Error(`Could not find CITIC transaction headers in ${source.file}`);
  const headers = rows[headerIndex].map(clean);
  const dateHeader = headers.find((value) => /交易日期/.test(value));
  const amountHeader = headers.find((value) => /交易金额/.test(value));
  const incomeHeader = headers.find((value) => /收入金额/.test(value));
  const expenseHeader = headers.find((value) => /支出金额/.test(value));
  const typeHeader = headers.find((value) => /交易类型|交易摘要|摘要/.test(value));
  const descriptionHeader = headers.find((value) => /交易对方|对方户名|备注/.test(value));
  const records = recordsFromRows(rows.slice(headerIndex), [dateHeader, amountHeader || incomeHeader, expenseHeader], source.file, headerIndex);
  return records.map((record) => {
    const income = numericOrNull(record[incomeHeader]);
    const expense = numericOrNull(record[expenseHeader]);
    const signedAmount = amountHeader ? parseAmount(record[amountHeader]) : income !== null ? income : -Math.abs(expense);
    return normalizedTransaction(source, record, {
    date: record[dateHeader],
    amount: signedAmount,
    type: record[typeHeader] || "Bank Transaction",
    description: record[descriptionHeader] || record[typeHeader] || "CITIC transaction",
    category: "",
    });
  });
}

function normalizedTransaction(source, record, values) {
  const date = parseDateText(values.date);
  if (!date) throw new Error(`Invalid transaction date in ${source.file} row ${record.__row}: ${values.date}`);
  return {
    source: source.file,
    sourceFolder: source.folder,
    sourceRow: record.__row,
    account: source.account,
    currency: source.currency,
    date,
    signedAmount: parseAmount(values.amount),
    type: clean(values.type),
    description: clean(values.description),
    rawCategory: clean(values.category),
    direction: clean(values.direction),
  };
}

function classifyTransaction(transaction, rule) {
  if (rule) {
    return {
      kind: rule.kind,
      category: rule.category,
      counterparty: rule.counterparty || rule.pair_id,
      description: rule.description || transaction.description,
      mirrorAccount: rule.mirror_account,
      fxRate: rule.fx_rate,
    };
  }

  if (/^Chase_Credit_\d{4}$/.test(transaction.sourceFolder)) {
    if (transaction.type === "Payment") return { kind: "ignore" };
    if (["Sale", "Return"].includes(transaction.type)) {
      return { kind: "expense", category: mapChaseCategory(transaction.rawCategory), description: transaction.description };
    }
  }

  if (/^Chase_Checking_\d{4}$/.test(transaction.sourceFolder)) {
    if (["LOAN_PMT"].includes(transaction.type) || /CHASE CREDIT CRD AUTOPAY/i.test(transaction.description)) return { kind: "ignore" };
    if (transaction.type === "DEBIT_CARD") return { kind: "expense", category: "", description: transaction.description };
    if (transaction.type === "ATM" && /ATM CASH DEPOSIT/i.test(transaction.description)) {
      return { kind: "transfer", category: "Internal Transfer", counterparty: `Cash → ${transaction.account}`, mirrorAccount: "Cash" };
    }
    if (transaction.type === "BILLPAY" && /Robinhood/i.test(transaction.description)) {
      return { kind: "transfer", category: "Internal Transfer", counterparty: `${transaction.account} → Robinhood Cash`, mirrorAccount: "Robinhood Cash" };
    }
    if (transaction.type === "ACH_DEBIT" && /SCHWAB BROKERAGE/i.test(transaction.description)) {
      return { kind: "transfer", category: "Internal Transfer", counterparty: `${transaction.account} → Charles Cash`, mirrorAccount: "Charles Cash" };
    }
    if (transaction.type === "MISC_CREDIT" && /FID BKG SVC/i.test(transaction.description)) {
      return { kind: "transfer", category: "Internal Transfer", counterparty: `Fidelity Cash → ${transaction.account}`, mirrorAccount: "Fidelity Cash" };
    }
  }

  if (transaction.sourceFolder === "Alipay" && ["收入", "支出"].includes(transaction.direction)) {
    return { kind: transaction.direction === "收入" ? "income" : "expense", category: "", description: transaction.description };
  }
  if (transaction.sourceFolder === "WeChat" && transaction.type === "商户消费") {
    return { kind: "expense", category: "", description: transaction.description };
  }
  return null;
}

function toCashflowItem(transaction, result) {
  const kind = result.kind;
  const amount = kind === "expense" ? -transaction.signedAmount : transaction.signedAmount;
  const description = withoutDisplayDate(result.description || transaction.description);
  const counterparty = withoutDisplayDate(result.counterparty);
  return {
    kind,
    date: transaction.date,
    source: transaction.source,
    sourceRow: transaction.sourceRow,
    fxRate: Number.isFinite(result.fxRate) ? result.fxRate : null,
    keyParts: kind === "transfer"
      ? [transaction.date, counterparty, result.category, transaction.type, description, amount, transaction.currency, transaction.account]
      : [transaction.date, result.category, transaction.type, description, amount, transaction.currency, transaction.account],
    values: kind === "transfer"
      ? [dateValue(transaction.date), transaction.date.slice(0, 7), counterparty, result.category || null, transaction.type || null, description || null, amount, transaction.currency, transaction.account]
      : [dateValue(transaction.date), transaction.date.slice(0, 7), result.category || null, transaction.type || null, description || null, amount, transaction.currency, transaction.account],
  };
}

function ruleMatches(rule, transaction) {
  if (rule.source_folder && rule.source_folder !== transaction.sourceFolder) return false;
  if (rule.date && rule.date !== transaction.date) return false;
  if (rule.amount !== null && Math.abs(rule.amount - transaction.signedAmount) > 0.005) return false;
  if (rule.type && rule.type !== transaction.type) return false;
  if (rule.description_contains && !transaction.description.toLowerCase().includes(rule.description_contains.toLowerCase())) return false;
  return Boolean(rule.source_folder || rule.date || rule.amount !== null || rule.type || rule.description_contains);
}

function mapChaseCategory(value) {
  return ({
    "Food & Drink": "Food",
    Groceries: "Grocery",
    Shopping: "Shopping",
    "Bills & Utilities": "Bills & Utilities",
    Travel: "Travel",
    "Health & Wellness": "Health & Wellness",
    Entertainment: "Entertainment",
    Automotive: "Transport",
  })[clean(value)] || clean(value);
}

function recordsFromRows(rows, requiredHeaders, file, rowOffset = 0) {
  if (!rows.length) throw new Error(`CashFlow source is empty: ${file}`);
  const headers = rows[0].map(clean);
  for (const header of requiredHeaders) if (header && !headers.includes(header)) throw new Error(`Missing header "${header}" in ${file}`);
  return rows.slice(1).filter(hasContent).map((row, index) => {
    const record = { __row: rowOffset + index + 2 };
    headers.forEach((header, column) => { if (header) record[header] = row[column] ?? ""; });
    return record;
  });
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = "";
  let quoted = false;
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    if (quoted) {
      if (character === '"' && text[index + 1] === '"') { field += '"'; index += 1; }
      else if (character === '"') quoted = false;
      else field += character;
    } else if (character === '"') quoted = true;
    else if (character === ",") { row.push(field); field = ""; }
    else if (character === "\n") { row.push(field.replace(/\r$/, "")); rows.push(row); row = []; field = ""; }
    else field += character;
  }
  if (field || row.length) { row.push(field.replace(/\r$/, "")); rows.push(row); }
  return rows;
}

function parseDateText(value) {
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "number") return new Date(Date.UTC(1899, 11, 30) + Math.floor(value) * 86400000).toISOString().slice(0, 10);
  const datePart = clean(value).split(/[T\s]/)[0];
  const compact = /^(\d{4})(\d{2})(\d{2})$/.exec(datePart);
  const iso = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(datePart);
  const us = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(datePart);
  if (!compact && !iso && !us) return null;
  const [year, month, day] = compact ? [compact[1], compact[2], compact[3]] : iso ? [iso[1], iso[2], iso[3]] : [us[3], us[1], us[2]];
  return `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
}

function dateValue(text) {
  return new Date(`${text}T00:00:00.000Z`);
}

function parseAmount(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const original = clean(value);
  const negative = /^\(.*\)$/.test(original);
  const normalized = original.replace(/[()$,￥¥元]/g, "").trim();
  const number = Number(normalized) * (negative ? -1 : 1);
  if (!Number.isFinite(number)) throw new Error(`Invalid transaction amount: ${value}`);
  return number;
}

function numericOrNull(value) {
  const text = clean(value);
  if (!text || text === "__" || text === "--") return null;
  return parseAmount(text);
}

function hasContent(row) {
  return row.some((value) => clean(value) !== "");
}

function clean(value) {
  return value === null || value === undefined ? "" : String(value).trim();
}

function withoutDisplayDate(value) {
  return clean(value)
    .replace(/\b(?:19|20)\d{2}[-/]\d{1,2}[-/]\d{1,2}\b/g, "")
    .replace(/\b\d{1,2}\/\d{1,2}\b/g, "")
    .replace(/\s{2,}/g, " ")
    .trim();
}

async function exists(file) {
  try {
    await fs.access(file);
    return true;
  } catch {
    return false;
  }
}

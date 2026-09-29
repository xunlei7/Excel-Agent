#!/usr/bin/env node

import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { continueTreasuryWorkbook, createTreasuryWorkbook } from "./create_workbook.mjs";
import { runCashflowImport } from "./import_cashflow.mjs";
import { runStockImport } from "./import_stock.mjs";
import { runAssetUpdate } from "./update_asset.mjs";
import { runSummaryUpdate } from "./update_summary.mjs";
import { runWorkbookFormat } from "./apply_workbook_format.mjs";

/**
 * Build a complete monthly-close workbook by composing the same four modules used
 * by the individual sheet tests. Intermediate files live in a temporary folder;
 * only the completed workbook is published.
 */
export async function runTreasuryUpdate({
  inputDir,
  inputFile,
  outputPath = null,
  cashflowCsvs = [],
  stockSources = [],
  throughDate = null,
  cashflowCacheDir = null,
  useCashflowCache = true,
  refreshCashflowCache = false,
  priceCsv = null,
  statementCacheDir = null,
  useStatementCache = true,
  refreshStatementCache = false,
  fetchStockPrices = true,
  dryRun = false,
  allowExistingOutput = false,
  mode = "start",
  bootstrapWorkbook = null,
}) {
  const root = path.resolve(inputDir);
  const closeInput = path.resolve(inputFile || path.join(root, "monthly-close-input.md"));
  const temporaryDirectory = await fs.mkdtemp(path.join(os.tmpdir(), "wealth-monthly-close-"));
  const stage = (name) => path.join(temporaryDirectory, name);
  let publishedOutput = null;

  try {
    if (!["start", "continue"].includes(mode)) fail(`Unsupported close mode: ${mode}. Use start or continue.`);
    const asOfDate = await readAsOfDate(closeInput);
    const cashflowCutoff = throughDate || asOfDate;
    const bootstrapSource = path.resolve(bootstrapWorkbook || path.join(root, "Output", "bootstrap", "One_Piece.xlsx"));
    const bootstrap = mode === "continue"
      ? await continueTreasuryWorkbook({ bootstrapWorkbookPath: bootstrapSource, outputPath: stage("00-bootstrap.xlsx") })
      : await createTreasuryWorkbook({ inputFile: closeInput, outputPath: stage("00-bootstrap.xlsx") });

    const cashflow = await runCashflowImport({
      workbookPath: bootstrap.output,
      outputPath: stage("10-cashflow.xlsx"),
      cashflowDir: path.join(root, "CashFlow"),
      cashflowCsvs,
      throughDate: cashflowCutoff,
      cacheDir: cashflowCacheDir,
      useCache: useCashflowCache,
      refreshCache: refreshCashflowCache,
    });
    const closeDate = latestCashflowDate(cashflow);
    const closeMonth = closeDate.slice(0, 7);
    publishedOutput = path.resolve(outputPath || path.join(root, "Output", closeMonth, `One_Piece_${closeMonth}.xlsx`));

    if (!dryRun && !allowExistingOutput && await exists(publishedOutput)) {
      fail(`Output already exists: ${publishedOutput}`);
    }

    const stock = await runStockImport({
      workbookPath: cashflow.output,
      outputPath: stage("20-stock.xlsx"),
      inputDir: root,
      inputFile: closeInput,
      stockSources,
      allowNewStockRows: true,
      fetchPrices: fetchStockPrices,
      priceAsOf: closeDate,
      cacheDir: statementCacheDir,
      useCache: useStatementCache,
      refreshCache: refreshStatementCache,
    });

    const asset = await runAssetUpdate({
      workbookPath: stock.output,
      outputPath: stage("30-asset.xlsx"),
      inputFile: closeInput,
      assetDate: closeDate,
    });

    const summary = await runSummaryUpdate({
      workbookPath: asset.output,
      outputPath: stage("40-summary.xlsx"),
      inputDir: root,
      inputFile: closeInput,
      priceCsv,
      summaryDate: closeDate,
    });

    const formatting = await runWorkbookFormat({
      workbookPath: summary.output,
      outputPath: stage("50-formatted.xlsx"),
    });

    if (!dryRun) {
      await fs.mkdir(path.dirname(publishedOutput), { recursive: true });
      await fs.copyFile(formatting.output, publishedOutput);
    }

    return {
      output: dryRun ? null : publishedOutput,
      plannedOutput: publishedOutput,
      closeDate,
      closeMonth,
      asOfDate,
      cashflowCutoff,
      dryRun,
      mode,
      stages: {
        bootstrap: publicBootstrapReport(bootstrap),
        cashflow: publicStageReport(cashflow),
        stock: publicStageReport(stock),
        asset: publicStageReport(asset),
        summary: publicStageReport(summary),
        formatting: publicStageReport(formatting),
      },
    };
  } finally {
    await fs.rm(temporaryDirectory, { recursive: true, force: true });
  }
}

function latestCashflowDate(report) {
  const dates = (report.files || [])
    .map((file) => file.latestDate)
    .filter((value) => /^\d{4}-\d{2}-\d{2}$/.test(String(value)))
    .sort();
  if (!dates.length) fail("No recognized CashFlow rows contain a valid date; the close month cannot be determined.");
  return dates.at(-1);
}

async function readAsOfDate(inputFile) {
  const text = await fs.readFile(inputFile, "utf8");
  const value = text.match(/^As-of Date:\s*(.*?)\s*$/mi)?.[1]?.trim() || "";
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    fail("monthly-close-input.md needs As-of Date as YYYY-MM-DD.");
  }
  return value;
}

function publicBootstrapReport(report) {
  const { output: _output, ...details } = report;
  return details;
}

function publicStageReport(report) {
  const { workbook: _workbook, output: _output, ...details } = report;
  return details;
}

function parseCliArgs(argv) {
  const args = {};
  const flags = new Set([
    "--dry-run",
    "--allow-existing-output",
    "--no-statement-cache",
    "--refresh-statement-cache",
    "--no-cashflow-cache",
    "--refresh-cashflow-cache",
    "--skip-stock-price-fetch",
  ]);
  for (let index = 0; index < argv.length; index += 1) {
    const key = argv[index];
    if (!key.startsWith("--")) fail(`Unexpected argument: ${key}`);
    const property = key.slice(2).replace(/-([a-z])/g, (_, character) => character.toUpperCase());
    if (flags.has(key)) args[property] = true;
    else {
      const value = argv[++index];
      if (!value || value.startsWith("--")) fail(`Missing value for ${key}.`);
      args[property] = value;
    }
  }
  return args;
}

function usage() {
  return [
    "Usage: update_treasury.mjs [options]",
    "",
    "Builds CashFlow, Stock, Asset, and Summary from code and raw inputs.",
    "",
    "Options:",
    "  --input-dir DIR                 Project root (default: current directory)",
    "  --input-file FILE               Monthly close input (default: monthly-close-input.md)",
    "  --mode start|continue            Start from code or continue from bootstrap (default: start)",
    "  --bootstrap-workbook FILE        Bootstrap source for continue mode",
    "  --output FILE                   Override Output/YYYY-MM/One_Piece_YYYY-MM.xlsx",
    "  --through-date YYYY-MM-DD        Ignore later CashFlow rows",
    "  --cashflow-cache-dir DIR         CashFlow parse-cache location",
    "  --no-cashflow-cache              Reparse every CashFlow source",
    "  --refresh-cashflow-cache         Refresh cached CashFlow parses",
    "  --price-csv FILE                 Use supplied Summary daily prices",
    "  --statement-cache-dir DIR        Brokerage statement cache location",
    "  --no-statement-cache             Reparse every brokerage statement",
    "  --refresh-statement-cache        Refresh cached statement parses",
    "  --skip-stock-price-fetch         Leave Stock current prices unrefreshed",
    "  --dry-run                        Complete all stages without publishing",
    "  --allow-existing-output          Replace an existing destination workbook",
  ].join("\n");
}

async function exists(file) {
  try { await fs.access(file); return true; } catch { return false; }
}

function fail(message) {
  throw new Error(message);
}

async function main() {
  const args = parseCliArgs(process.argv.slice(2));
  const inputDir = path.resolve(args.inputDir || process.cwd());
  const report = await runTreasuryUpdate({
    inputDir,
    inputFile: args.inputFile || path.join(inputDir, "monthly-close-input.md"),
    outputPath: args.output || null,
    throughDate: args.throughDate || null,
    cashflowCacheDir: args.cashflowCacheDir || null,
    useCashflowCache: !args.noCashflowCache,
    refreshCashflowCache: Boolean(args.refreshCashflowCache),
    priceCsv: args.priceCsv || null,
    statementCacheDir: args.statementCacheDir || null,
    useStatementCache: !args.noStatementCache,
    refreshStatementCache: Boolean(args.refreshStatementCache),
    fetchStockPrices: !args.skipStockPriceFetch,
    dryRun: Boolean(args.dryRun),
    allowExistingOutput: Boolean(args.allowExistingOutput),
    mode: args.mode || "start",
    bootstrapWorkbook: args.bootstrapWorkbook || null,
  });
  console.log(JSON.stringify(report, null, 2));
}

const directRun = process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href;
if (directRun) main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  console.error("\n" + usage());
  process.exitCode = 1;
});

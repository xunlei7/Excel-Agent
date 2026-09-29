import fs from "node:fs/promises";
import path from "node:path";
import { createRequire } from "node:module";

const dependencies = process.env.CODEX_NODE_MODULES;
if (!dependencies) throw new Error("CODEX_NODE_MODULES is required");
const require = createRequire(path.join(dependencies, "package.json"));
const { SpreadsheetFile } = require("@oai/artifact-tool");
const JSZip = require("jszip");

// artifact-tool may re-export imported shared formulas incompletely. Convert every
// formula represented in its workbook model into a normal per-cell OOXML formula.
export async function saveWorkbookFormulaSafe(workbook, workbookPath, { restoreFixedText = true } = {}) {
  const expectedFormulas = collectWorkbookFormulas(workbook);
  const exported = await SpreadsheetFile.exportXlsx(workbook);
  const archive = await JSZip.loadAsync(Buffer.from(exported.data));
  const workbookXml = await archive.file("xl/workbook.xml")?.async("string");
  const relationshipsXml = await archive.file("xl/_rels/workbook.xml.rels")?.async("string");
  if (!workbookXml || !relationshipsXml) fail("Exported workbook is missing its workbook metadata.");
  archive.file("xl/workbook.xml", forceAutomaticRecalculation(workbookXml));

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
    if (restoreFixedText) {
      worksheetXml = restoreFixedTextForSheet(worksheetXml, name);
    }
    archive.file(worksheetPath, worksheetXml);
  }
  await fs.writeFile(workbookPath, await archive.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

function forceAutomaticRecalculation(workbookXml) {
  const existing = /<([A-Za-z_][\w.-]*:)?calcPr\b[^>]*\/?\s*>/;
  if (existing.test(workbookXml)) {
    return workbookXml.replace(existing, (_match, prefix = "") => `<${prefix}calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>`);
  }
  return workbookXml.replace(/<\/([A-Za-z_][\w.-]*:)?workbook>/, (_match, prefix = "") => `<${prefix}calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></${prefix}workbook>`);
}

export function restoreFixedTextForSheet(worksheetXml, sheetName) {
  const fixedText = sheetName === "CashFlow" ? [
    ["A3", "Total Income (Currency)"],
    ["A4", "Total Expense (Currency)"],
    ["A5", "Total Transfer (Currency)"],
    ["A7", "Net Cash Flow"],
    ["A11", "Date"],
    ["B11", "Year-Month"],
    ["C11", "Category"],
    ["D11", "Type"],
    ["E11", "Description"],
    ["F11", "Amount"],
    ["G11", "Currency"],
    ["H11", "Account"],
    ["I11", "FX Rate"],
    ["J11", "Converted Amount"],
    ["N11", "Date"],
    ["O11", "Year-Month"],
    ["P11", "Category"],
    ["Q11", "Type"],
    ["R11", "Description"],
    ["S11", "Amount"],
    ["T11", "Currency"],
    ["U11", "Account"],
    ["V11", "FX Rate"],
    ["W11", "Converted Amount"],
    ["AA11", "Date"],
    ["AB11", "Year-Month"],
    ["AC11", "Counterparty"],
    ["AD11", "Category"],
    ["AE11", "Type"],
    ["AF11", "Description"],
    ["AG11", "Amount"],
    ["AH11", "Currency"],
    ["AI11", "Account"],
    ["AJ11", "FX Rate"],
    ["AK11", "Converted Amount"],
    ["AL11", "Outstanding"],
    ["AM11", "Status"],
  ] : sheetName === "Summary" ? [
    ["N4", "Real Asset"],
    ["N5", "Calculated Asset"],
    ["N6", "Total Difference"],
    ["N7", "Status"],
  ] : [];
  let result = worksheetXml;
  for (const [address, value] of fixedText) {
    const cellPattern = new RegExp(`<([A-Za-z_][\\w.-]*:)?c\\b([^>]*\\br="${address}"[^>]*)(?:\\/>|>[\\s\\S]*?<\\/\\1c>)`);
    let replaced = false;
    result = result.replace(cellPattern, (_cellXml, prefix = "", opening = "") => {
      replaced = true;
      const attributes = opening.replace(/\s+t="[^"]*"/g, "").replace(/\s*\/$/, "");
      return `<${prefix}c${attributes} t="str"><${prefix}v>${escapeXml(value)}</${prefix}v></${prefix}c>`;
    });
    if (!replaced) result = insertFixedTextCell(result, address, value);
  }
  return result;
}

function insertFixedTextCell(worksheetXml, address, value) {
  const rowNumber = address.match(/\d+$/)?.[0];
  const targetColumn = columnIndex(address.match(/^[A-Z]+/)?.[0] || "A");
  if (!rowNumber) return worksheetXml;
  const rowPattern = new RegExp(`<([A-Za-z_][\\w.-]*:)?row\\b([^>]*\\br="${rowNumber}"[^>]*)>([\\s\\S]*?)<\\/\\1row>`);
  return worksheetXml.replace(rowPattern, (rowXml, prefix = "", rowAttributes = "", body = "") => {
    const cells = [...body.matchAll(/<([A-Za-z_][\w.-]*:)?c\b[^>]*\br="([A-Z]+)\d+"[^>]*(?:\/>|>[\s\S]*?<\/\1c>)/g)];
    const nextCell = cells.find((match) => columnIndex(match[2]) > targetColumn);
    const neighbor = nextCell || cells.at(-1);
    const style = neighbor?.[0].match(/\bs="([^"]+)"/)?.[1];
    const styleAttribute = style ? ` s="${style}"` : "";
    const newCell = `<${prefix}c r="${address}"${styleAttribute} t="str"><${prefix}v>${escapeXml(value)}</${prefix}v></${prefix}c>`;
    const insertionIndex = nextCell?.index ?? body.length;
    const updatedBody = `${body.slice(0, insertionIndex)}${newCell}${body.slice(insertionIndex)}`;
    return `<${prefix}row${rowAttributes}>${updatedBody}</${prefix}row>`;
  });
}

function collectWorkbookFormulas(workbook) {
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
      formulaMap.set(`${columnLetter(start.col + colOffset)}${start.row + rowOffset}`, formula);
    }));
    result.set(sheet.name, formulaMap);
  }
  return result;
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

function parseCellAddress(address) {
  const match = /^([A-Z]+)(\d+)$/.exec(address.toUpperCase());
  if (!match) fail(`Invalid cell address: ${address}`);
  return { col: columnIndex(match[1]), row: Number(match[2]) };
}

function columnIndex(letters) {
  let result = 0;
  for (const character of letters) result = result * 26 + character.charCodeAt(0) - 64;
  return result - 1;
}

function columnLetter(index) {
  let value = index + 1;
  let result = "";
  while (value > 0) {
    value -= 1;
    result = String.fromCharCode(65 + (value % 26)) + result;
    value = Math.floor(value / 26);
  }
  return result;
}

function escapeXml(value) {
  return String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function fail(message) { throw new Error(message); }

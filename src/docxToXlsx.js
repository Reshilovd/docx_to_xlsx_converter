import JSZip from "jszip";
import ExcelJS from "exceljs";

const getName = (n) => (n?.localName || n?.nodeName || "").split(":").pop();
const direct = (n, tag) =>
  Array.from(n?.childNodes || []).filter((c) => c.nodeType === Node.ELEMENT_NODE && getName(c) === tag);
const first = (n, tag) => direct(n, tag)[0] || null;
const all = (n, tag) =>
  Array.from(n?.getElementsByTagName("*") || []).filter((c) => getName(c) === tag);

function normalizeQuestionCode(rawCode) {
  return rawCode.replace(/\s+/g, "").toUpperCase();
}

function extractQuestionCode(text) {
  const match = text.trim().match(/^([A-Za-zА-Яа-я]\d+[A-Za-zА-Яа-я]?|\d+)\s*[\.\)]/);
  return match ? normalizeQuestionCode(match[1]) : null;
}

function toArgb(raw) {
  const v = String(raw || "").trim().replace(/^#/, "");
  if (/^[0-9a-f]{3}$/i.test(v)) {
    const x = v
      .split("")
      .map((c) => c + c)
      .join("")
      .toUpperCase();
    return `FF${x}`;
  }
  if (/^[0-9a-f]{6}$/i.test(v)) {
    return `FF${v.toUpperCase()}`;
  }
  return null;
}

function styleFromRPr(rPr) {
  const font = {};
  if (!rPr) return font;
  if (first(rPr, "b")) font.bold = true;
  if (first(rPr, "i")) font.italic = true;
  const u = first(rPr, "u");
  if (u) {
    const val = (u.getAttribute("w:val") || u.getAttribute("val") || "single").toLowerCase();
    if (val !== "none") font.underline = true;
  }
  if (first(rPr, "strike") || first(rPr, "dstrike")) font.strike = true;
  const color = first(rPr, "color");
  if (color) {
    const argb = toArgb(color.getAttribute("w:val") || color.getAttribute("val"));
    if (argb) font.color = { argb };
  }
  const size = first(rPr, "sz");
  if (size) {
    const val = Number(size.getAttribute("w:val") || size.getAttribute("val"));
    if (Number.isFinite(val) && val > 0) font.size = val / 2;
  }
  const rFonts = first(rPr, "rFonts");
  if (rFonts) {
    const name =
      rFonts.getAttribute("w:ascii") ||
      rFonts.getAttribute("ascii") ||
      rFonts.getAttribute("w:hAnsi") ||
      rFonts.getAttribute("hAnsi");
    if (name) font.name = name;
  }
  const vert = first(rPr, "vertAlign");
  if (vert) {
    const val = (vert.getAttribute("w:val") || vert.getAttribute("val") || "").toLowerCase();
    if (val === "superscript") font.vertAlign = "superscript";
    if (val === "subscript") font.vertAlign = "subscript";
  }
  return font;
}

function parseRun(runNode) {
  const font = styleFromRPr(first(runNode, "rPr"));
  const parts = [];
  Array.from(runNode.childNodes || []).forEach((child) => {
    if (child.nodeType !== Node.ELEMENT_NODE) return;
    const tag = getName(child);
    if (tag === "t" || tag === "instrText") parts.push(child.textContent || "");
    else if (tag === "tab") parts.push("\t");
    else if (tag === "br" || tag === "cr") parts.push("\n");
  });
  const text = parts.join("");
  return text ? { text, font } : null;
}

function mergeRuns(runs) {
  const out = [];
  runs.forEach((run) => {
    if (!run?.text) return;
    const prev = out[out.length - 1];
    if (prev && JSON.stringify(prev.font || {}) === JSON.stringify(run.font || {})) {
      prev.text += run.text;
    } else {
      out.push({ text: run.text, font: { ...(run.font || {}) } });
    }
  });
  return out;
}

function hasFormat(font) {
  return Boolean(
    font.bold ||
      font.italic ||
      font.underline ||
      font.strike ||
      font.vertAlign ||
      font.name ||
      font.size ||
      (font.color && font.color.argb)
  );
}

function runsToCell(runs) {
  const compact = mergeRuns(runs);
  if (compact.length === 0) return "";
  if (compact.length === 1 && !hasFormat(compact[0].font || {})) return compact[0].text;
  return { richText: compact };
}

function splitCellLines(cellValue) {
  if (typeof cellValue === "string") {
    return cellValue.split(/\n+/).map((x) => x.trim()).filter(Boolean);
  }
  if (!cellValue?.richText) return [];
  const lines = [];
  let current = [];
  cellValue.richText.forEach((run) => {
    String(run.text || "")
      .split("\n")
      .forEach((part, i, arr) => {
        if (part) current.push({ text: part, font: { ...(run.font || {}) } });
        if (i < arr.length - 1) {
          if (current.length > 0) lines.push({ richText: current });
          current = [];
        }
      });
  });
  if (current.length > 0) lines.push({ richText: current });
  return lines.filter((line) =>
    line.richText.some((run) => String(run.text || "").trim() !== "")
  );
}

function parseParagraph(pNode) {
  const runs = all(pNode, "r").map(parseRun).filter(Boolean);
  return runsToCell(runs);
}

function joinCells(values) {
  const runs = [];
  values.forEach((v, idx) => {
    if (idx > 0) runs.push({ text: " | ", font: {} });
    if (typeof v === "string") runs.push({ text: v, font: {} });
    else if (v?.richText) runs.push(...v.richText.map((r) => ({ text: r.text, font: { ...(r.font || {}) } })));
  });
  return runsToCell(runs);
}

function parseTable(tblNode) {
  const lines = [];
  direct(tblNode, "tr").forEach((row) => {
    const values = direct(row, "tc")
      .map((tc) => {
        const runs = [];
        direct(tc, "p").forEach((p, i) => {
          const value = parseParagraph(p);
          if (!value) return;
          if (i > 0) runs.push({ text: "\n", font: {} });
          if (typeof value === "string") runs.push({ text: value, font: {} });
          else runs.push(...value.richText);
        });
        return runsToCell(runs);
      })
      .filter((v) => (typeof v === "string" ? v.trim() !== "" : v?.richText?.length > 0));
    if (values.length > 0) lines.push(...splitCellLines(joinCells(values)));
  });
  return lines;
}

function cellToText(cellValue) {
  if (typeof cellValue === "string") return cellValue.trim();
  return (cellValue?.richText || []).map((r) => String(r.text || "")).join("").trim();
}

function isAuxiliaryQuestionLine(text) {
  const normalized = text.trim();
  if (!normalized) return false;

  // Instructions like "(choose one answer)" should not shift answer alignment.
  if (/^\([^()]{3,}\)$/.test(normalized)) return true;

  // Section titles are often embedded between questions and should not shift answers.
  if (/^(блок|block)\b/i.test(normalized)) return true;

  return false;
}

function buildAlignedRecords(lines) {
  const records = [];
  const occurrences = new Map();
  let currentCode = null;
  let currentOcc = 0;
  let lineIdx = 0;
  let auxIdx = 0;
  let preface = 0;

  lines.forEach((line) => {
    const text = cellToText(line);
    if (!text) return;
    const code = extractQuestionCode(text);
    if (code) {
      const occ = (occurrences.get(code) || 0) + 1;
      occurrences.set(code, occ);
      currentCode = code;
      currentOcc = occ;
      lineIdx = 0;
      auxIdx = 0;
      records.push({ key: `Q:${code}:${occ}:0`, value: line });
      return;
    }
    if (currentCode) {
      if (isAuxiliaryQuestionLine(text)) {
        auxIdx += 1;
        records.push({ key: `Q:${currentCode}:${currentOcc}:AUX:${auxIdx}`, value: line });
        return;
      }
      lineIdx += 1;
      records.push({ key: `Q:${currentCode}:${currentOcc}:${lineIdx}`, value: line });
      return;
    }
    preface += 1;
    records.push({ key: `P:${preface}`, value: line });
  });

  return records;
}

function cellValueToRuns(value) {
  if (value == null) return [];
  if (typeof value === "string") return [{ text: value, font: {} }];
  if (typeof value === "number" || typeof value === "boolean") return [{ text: String(value), font: {} }];
  if (value instanceof Date) return [{ text: value.toISOString(), font: {} }];
  if (value.richText) {
    return value.richText
      .map((run) => ({ text: String(run.text || ""), font: { ...(run.font || {}) } }))
      .filter((run) => run.text !== "");
  }
  if (value.text) return [{ text: String(value.text), font: {} }];
  if (value.result != null) return [{ text: String(value.result), font: {} }];
  return [{ text: String(value), font: {} }];
}

function escapeHtmlAttr(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function wrapRunsWithLangSpan(runs, className) {
  const safeClass = escapeHtmlAttr(className || "lang");
  return [
    { text: `<span class="${safeClass}">`, font: {} },
    ...runs,
    { text: "</span>", font: {} },
  ];
}

function buildSpanLine(value) {
  const runs = cellValueToRuns(value);
  if (runs.length === 0) return "";
  return runsToCell(runs);
}

function isEmptyCellValue(value) {
  return cellToText(value) === "";
}

function autoFit(worksheet) {
  worksheet.columns.forEach((column) => {
    let max = 8;
    column.eachCell({ includeEmpty: true }, (cell) => {
      let text = "";
      if (typeof cell.value === "string") text = cell.value;
      else if (cell.value?.richText) text = cell.value.richText.map((r) => r.text).join("");
      const length = String(text).split("\n").reduce((m, part) => Math.max(m, part.length), 0);
      max = Math.max(max, Math.min(length + 2, 80));
    });
    column.width = max;
  });
}

async function extractDocumentRecords(file) {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const xmlFile = zip.file("word/document.xml");
  if (!xmlFile) return [{ key: "P:1", value: "Document is empty." }];

  const xml = new DOMParser().parseFromString(await xmlFile.async("text"), "application/xml");
  const body = all(xml, "body")[0];
  if (!body) return [{ key: "P:1", value: "Document is empty." }];

  const lines = [];
  Array.from(body.childNodes || []).forEach((node) => {
    if (node.nodeType !== Node.ELEMENT_NODE) return;
    const tag = getName(node);
    if (tag === "p") lines.push(...splitCellLines(parseParagraph(node)));
    if (tag === "tbl") {
      const tableLines = parseTable(node);
      if (tableLines.length > 0) {
        lines.push(...tableLines);
        lines.push("");
      }
    }
  });

  if (lines.length === 0) return [{ key: "P:1", value: "Document is empty." }];
  return buildAlignedRecords(lines);
}

export async function convertDocxFileToXlsxBlob(file) {
  return convertDocxFilesToXlsxBlob([file], { mode: "columns", classNames: ["default"] });
}

export async function convertDocxFilesToXlsxBlob(files, options = []) {
  const parsed = Array.isArray(options)
    ? { mode: "spans", classNames: options }
    : { mode: options.mode || "spans", classNames: options.classNames || [] };
  const mode = parsed.mode === "columns" ? "columns" : "spans";

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("content");
  const docs = await Promise.all(files.map((file) => extractDocumentRecords(file)));
  const maps = docs.map((records) => {
    const map = new Map();
    records.forEach((record) => map.set(record.key, record.value));
    return map;
  });

  const keys = [];
  const seen = new Set();
  docs.forEach((records) => {
    records.forEach((record) => {
      if (!seen.has(record.key)) {
        seen.add(record.key);
        keys.push(record.key);
      }
    });
  });

  keys.forEach((key) => {
    if (mode === "columns") {
      const row = sheet.addRow(maps.map((map) => map.get(key) || ""));
      row.eachCell((cell) => {
        cell.alignment = { vertical: "top", wrapText: true };
      });
      return;
    }
    const mergedRuns = maps
      .map((map, index) => {
        const value = map.get(key);
        if (!value || isEmptyCellValue(value)) return [];
        const rawRuns = cellValueToRuns(value);
        return wrapRunsWithLangSpan(rawRuns, parsed.classNames[index] || `lang_${index + 1}`);
      })
      .flat();
    const row = sheet.addRow([runsToCell(mergedRuns)]);
    row.getCell(1).alignment = { vertical: "top", wrapText: true };
  });

  if (keys.length === 0) {
    const row = sheet.addRow(["Document is empty."]);
    row.getCell(1).alignment = { vertical: "top", wrapText: true };
  }

  autoFit(sheet);
  const bytes = await workbook.xlsx.writeBuffer();
  return new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

export async function inspectXlsxColumns(file) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  const sheet = workbook.worksheets[0];
  if (!sheet) return 0;

  let columnCount = 0;
  sheet.eachRow({ includeEmpty: false }, (row) => {
    columnCount = Math.max(columnCount, row.actualCellCount || 0);
  });

  return columnCount;
}

export async function convertCheckedXlsxToSpansBlob(file, classNames = []) {
  const srcWorkbook = new ExcelJS.Workbook();
  await srcWorkbook.xlsx.load(await file.arrayBuffer());
  const srcSheet = srcWorkbook.worksheets[0];
  if (!srcSheet) {
    throw new Error("В XLSX не найден лист с данными.");
  }

  const maxColumns = await inspectXlsxColumns(file);
  if (maxColumns === 0) {
    throw new Error("В XLSX нет данных для объединения.");
  }

  const outWorkbook = new ExcelJS.Workbook();
  const outSheet = outWorkbook.addWorksheet("content");

  srcSheet.eachRow({ includeEmpty: false }, (row) => {
    const mergedRuns = [];
    for (let i = 1; i <= maxColumns; i += 1) {
      const rawRuns = cellValueToRuns(row.getCell(i).value);
      if (rawRuns.length === 0 || rawRuns.every((run) => run.text.trim() === "")) continue;
      mergedRuns.push(
        ...wrapRunsWithLangSpan(rawRuns, classNames[i - 1] || `lang_${i}`)
      );
    }
    outSheet.addRow([runsToCell(mergedRuns)]);
  });

  autoFit(outSheet);
  const bytes = await outWorkbook.xlsx.writeBuffer();
  return new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

export const __testables = {
  normalizeQuestionCode,
  extractQuestionCode,
  isAuxiliaryQuestionLine,
  buildAlignedRecords,
  cellValueToRuns,
  wrapRunsWithLangSpan,
  runsToCell,
};


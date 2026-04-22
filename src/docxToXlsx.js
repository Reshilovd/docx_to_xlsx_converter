import mammoth from "mammoth";
import ExcelJS from "exceljs";

function normalizeWhitespace(text) {
  return text.replace(/\s+/g, " ");
}

function buildRichTextRuns(node, style = { bold: false, italic: false, underline: false }) {
  if (node.nodeType === Node.TEXT_NODE) {
    const text = normalizeWhitespace(node.textContent || "");
    if (!text.trim()) {
      return [];
    }

    return [
      {
        text,
        font: { ...style },
      },
    ];
  }

  if (node.nodeType !== Node.ELEMENT_NODE) {
    return [];
  }

  const tag = node.tagName.toLowerCase();
  if (tag === "br") {
    return [{ text: "\n", font: { ...style } }];
  }

  const nextStyle = {
    bold: style.bold || tag === "strong" || tag === "b",
    italic: style.italic || tag === "em" || tag === "i",
    underline: style.underline || tag === "u",
  };

  const childRuns = [];
  Array.from(node.childNodes).forEach((child) => {
    childRuns.push(...buildRichTextRuns(child, nextStyle));
  });

  if (["p", "div", "li"].includes(tag)) {
    childRuns.push({ text: "\n", font: { ...style } });
  }

  return childRuns;
}

function cleanupRuns(runs) {
  const compact = [];
  runs.forEach((run) => {
    if (!run.text) {
      return;
    }
    const previous = compact[compact.length - 1];
    if (
      previous &&
      previous.font.bold === run.font.bold &&
      previous.font.italic === run.font.italic &&
      previous.font.underline === run.font.underline
    ) {
      previous.text += run.text;
    } else {
      compact.push(run);
    }
  });

  while (compact.length && compact[0].text === "\n") {
    compact.shift();
  }
  while (compact.length && compact[compact.length - 1].text === "\n") {
    compact.pop();
  }
  return compact;
}

function elementToCellValue(element) {
  const runs = cleanupRuns(buildRichTextRuns(element));
  if (runs.length === 0) {
    return "";
  }
  if (runs.length === 1 && !runs[0].font.bold && !runs[0].font.italic && !runs[0].font.underline) {
    return runs[0].text;
  }
  return { richText: runs };
}

function splitCellValueByNewLines(cellValue) {
  if (typeof cellValue === "string") {
    return cellValue
      .split(/\n+/)
      .map((part) => part.trim())
      .filter(Boolean);
  }

  if (!cellValue || !Array.isArray(cellValue.richText)) {
    return [];
  }

  const lines = [];
  let currentRuns = [];

  cellValue.richText.forEach((run) => {
    const parts = String(run.text || "").split("\n");
    parts.forEach((part, index) => {
      if (part) {
        currentRuns.push({ text: part, font: { ...run.font } });
      }
      if (index < parts.length - 1) {
        if (currentRuns.length > 0) {
          lines.push({ richText: currentRuns });
        }
        currentRuns = [];
      }
    });
  });

  if (currentRuns.length > 0) {
    lines.push({ richText: currentRuns });
  }

  return lines.filter((line) => {
    if (typeof line === "string") {
      return line.trim() !== "";
    }
    return line.richText.some((run) => String(run.text || "").trim() !== "");
  });
}

function htmlTableToRows(tableElement) {
  const rows = [];
  const trList = Array.from(tableElement.querySelectorAll("tr"));
  trList.forEach((tr) => {
    const cells = Array.from(tr.querySelectorAll("th,td"));
    rows.push(cells.map(elementToCellValue));
  });
  return rows.filter((row) =>
    row.some((cell) => {
      if (!cell) {
        return false;
      }
      if (typeof cell === "string") {
        return cell.trim() !== "";
      }
      return Array.isArray(cell.richText) && cell.richText.some((run) => run.text.trim() !== "");
    })
  );
}

function autoFitColumns(worksheet) {
  worksheet.columns.forEach((column) => {
    let maxLength = 8;
    column.eachCell({ includeEmpty: true }, (cell) => {
      let value = "";
      if (typeof cell.value === "string") {
        value = cell.value;
      } else if (cell.value && typeof cell.value === "object" && Array.isArray(cell.value.richText)) {
        value = cell.value.richText.map((r) => r.text).join("");
      }
      const cellLength = String(value).split("\n").reduce((max, part) => Math.max(max, part.length), 0);
      maxLength = Math.max(maxLength, Math.min(cellLength + 2, 80));
    });
    column.width = maxLength;
  });
}

export async function convertDocxFileToXlsxBlob(file) {
  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.convertToHtml(
    { arrayBuffer },
    {
      includeDefaultStyleMap: true,
      ignoreEmptyParagraphs: false,
    }
  );

  const parser = new DOMParser();
  const doc = parser.parseFromString(result.value, "text/html");
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("content");
  const allRows = [];

  const elements = Array.from(doc.body.children);
  elements.forEach((element) => {
    if (element.tagName.toLowerCase() === "table") {
      const tableRows = htmlTableToRows(element);
      if (tableRows.length > 0) {
        allRows.push(...tableRows);
        allRows.push([]);
      }
      return;
    }

    const cellValue = elementToCellValue(element);
    const lines = splitCellValueByNewLines(cellValue);
    lines.forEach((line) => allRows.push([line]));
  });

  const rowsToWrite = allRows.length > 0 ? allRows : [["Document is empty."]];
  rowsToWrite.forEach((row) => {
    const xlsxRow = worksheet.addRow(row);
    xlsxRow.eachCell((cell) => {
      cell.alignment = { vertical: "top", wrapText: true };
    });
  });
  autoFitColumns(worksheet);

  const xlsxBuffer = await workbook.xlsx.writeBuffer();
  return new Blob([xlsxBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

import { __testables } from "./docxToXlsx";

describe("docxToXlsx helpers", () => {
  test("cellValueToRuns returns empty array for nullish values", () => {
    expect(__testables.cellValueToRuns(null)).toEqual([]);
    expect(__testables.cellValueToRuns(undefined)).toEqual([]);
  });

  test("auxiliary line does not shift answer indexes", () => {
    const records = __testables.buildAlignedRecords([
      "A1. Question",
      "(Пожалуйста, выберите один вариант ответа.)",
      "Option one",
      "Option two",
    ]);
    expect(records.map((x) => x.key)).toEqual([
      "Q:A1:1:0",
      "Q:A1:1:AUX:1",
      "Q:A1:1:1",
      "Q:A1:1:2",
    ]);
  });

  test("wrapRunsWithLangSpan injects class wrapper tags", () => {
    const wrapped = __testables.wrapRunsWithLangSpan([{ text: "Привет", font: { bold: true } }], "ru");
    expect(wrapped).toHaveLength(3);
    expect(wrapped[0].text).toBe('<span class="ru">');
    expect(wrapped[1]).toEqual({ text: "Привет", font: { bold: true } });
    expect(wrapped[2].text).toBe("</span>");
  });

  test("combineCodeAndLabelCells joins code with label in one cell", () => {
    const merged = __testables.combineCodeAndLabelCells("Астана", "3");
    expect(__testables.cellToText(merged)).toBe("3 Астана");
  });

  test("dedupeRowValues removes repeated values in table row", () => {
    const deduped = __testables.dedupeRowValues(["7 DAYS", "1", "1", "1"]);
    expect(deduped.map((v) => __testables.cellToText(v))).toEqual(["7 DAYS", "1"]);
  });

  test("mergeCodeLabelArrays joins parallel label and code rows", () => {
    const merged = __testables.mergeCodeLabelArrays(
      ["Совсем не нравится", "Скорее не нравится", "Ни нравится, ни не нравится"],
      ["1", "2", "3"]
    );
    expect(merged.map((v) => __testables.cellToText(v))).toEqual([
      "1 Совсем не нравится",
      "2 Скорее не нравится",
      "3 Ни нравится, ни не нравится",
    ]);
  });

  test("mergeAdjacentCodeLabelLines joins vertical code-label pairs", () => {
    const merged = __testables.mergeAdjacentCodeLabelLines([
      "1",
      "Еуропадағы №1 круассан",
      "2",
      "Жүрек жалғау үшін мінсіз",
      "Q1. Некий вопрос",
      "3",
      "Следующий лейбл",
    ]);
    expect(merged.map((v) => __testables.cellToText(v))).toEqual([
      "1 Еуропадағы №1 круассан",
      "2 Жүрек жалғау үшін мінсіз",
      "Q1. Некий вопрос",
      "3 Следующий лейбл",
    ]);
  });

  test("mergeAdjacentCodeLabelLines does not shift when label already has inline code", () => {
    const merged = __testables.mergeAdjacentCodeLabelLines([
      "Мүлдем сенімге лайық емес",
      "Сенімен лайық емес сияқты",
      "Жауап беру қиын",
      "1 - Толығымен сенімділік береді",
      "2",
      "3",
      "4",
      "5",
    ]);
    expect(merged.map((v) => __testables.cellToText(v))).toEqual([
      "Мүлдем сенімге лайық емес",
      "Сенімен лайық емес сияқты",
      "Жауап беру қиын",
      "1 - Толығымен сенімділік береді",
      "2",
      "3",
      "4",
      "5",
    ]);
  });

  test("detectStructuredCodeLabelColumns detects code in first table column", () => {
    const result = __testables.detectStructuredCodeLabelColumns([
      ["1", "Алматы", "СВЕРИТЬ КВОТУ"],
      ["2", "Астана", "СВЕРИТЬ КВОТУ"],
      ["3", "Шымкент", "ЗАВЕРШИТЬ"],
    ]);
    expect(result).toEqual({ enabled: true, codeColumnIndex: 0 });
  });
});

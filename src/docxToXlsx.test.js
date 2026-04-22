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
});

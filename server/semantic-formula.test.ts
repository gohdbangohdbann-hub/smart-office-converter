import { describe, expect, it } from "vitest";
import { inferSafeFormulas, updateSuggestionStatus, acceptedFormulas } from "@shared/formula-inference";
import { buildSemanticTable } from "@shared/semantic";
import type { ExcelTablePlan } from "@shared/excel";

function table(): ExcelTablePlan {
  const values = [["الوصف", "الكمية", "سعر الوحدة", "المجموع الفرعي"], ["ورق", 2, 10, null], ["قلم", 3, 5, null]] as const;
  const cells = values.flatMap((row, rowIndex) => row.map((value, columnIndex) => ({ text: value == null ? "" : String(value), value, rowIndex, columnIndex, type: typeof value === "number" ? "number" as const : "text" as const, language: "ar" as const, format: { horizontalAlignment: "right" as const, wrapText: true, bold: rowIndex === 0, border: "thin" as const } })));
  return { name: "NawaTable1", sheetName: "Nawa OCR", rowCount: 3, columnCount: 4, headers: values[0].map(String), rows: [], columns: [], cells, values: values.map((row) => [...row]), direction: "rtl", mergeRanges: [] };
}

describe("Semantic Understanding and safe Formula Inference", () => {
  it("classifies Arabic headers and preserves source metadata", () => {
    const result = buildSemanticTable(table());
    expect(result.byType.quantity?.[0]).toMatchObject({ type: "quantity", confidence: 0.9, source: { rowIndex: 0, columnIndex: 1 } });
    expect(result.byType["unit-price"]?.length).toBe(3);
    expect(result.byType.description?.[1]?.source.sourceText).toBe("ورق");
  });
  it("suggests formulas only for numeric quantity and unit price cells", () => {
    const result = inferSafeFormulas(table());
    expect(result.suggestions.map((item) => item.formula)).toEqual(["=B2*C2", "=B3*C3"]);
    expect(result.suggestions.every((item) => item.status === "suggested")).toBe(true);
    expect(acceptedFormulas(result)).toHaveLength(0);
  });
  it("requires explicit review and rejects unsafe edited formulas", () => {
    const suggestion = inferSafeFormulas(table()).suggestions[0]!;
    expect(() => updateSuggestionStatus(suggestion, "edited", "B2*C2")).toThrow("FORMULA_MUST_START_WITH_EQUALS");
    const accepted = updateSuggestionStatus(suggestion, "accepted");
    expect(acceptedFormulas({ ...inferSafeFormulas(table()), suggestions: [accepted] })).toHaveLength(1);
    expect(updateSuggestionStatus(suggestion, "rejected").status).toBe("rejected");
  });
  it("does not invent formulas when semantic input is incomplete", () => {
    const incomplete = { ...table(), headers: ["الوصف", "الكمية", "ملاحظات", "المجموع الفرعي"], cells: table().cells.map((cell) => cell.columnIndex === 2 ? { ...cell, value: cell.rowIndex === 0 ? cell.value : "غير متاح", type: "text" as const } : cell) };
    expect(inferSafeFormulas(incomplete).suggestions).toHaveLength(0);
  });
});

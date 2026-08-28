import { describe, expect, it } from "vitest";
import { a1Address, buildExcelPlan, classifyExcelCell, resolveExcelOperation } from "@shared/excel";
import type { OCRDocument } from "@shared/ocr";
import { buildExcelTransformRequest, submitExcelTransform } from "@shared/excel-request";

const fixture: OCRDocument = {
  id: "excel-fixture", fileName: "فاتورة-عربية.pdf", mimeType: "application/pdf", pageCount: 1, language: "mixed", provider: "fixture", processedAt: new Date().toISOString(),
  privacy: { originalRetained: false, temporaryDataDeleted: true }, tables: [{ rowCount: 3, columnCount: 4, confidence: 0.91, cells: [
    { text: "رقم", rowIndex: 0, columnIndex: 0 }, { text: "المنتج", rowIndex: 0, columnIndex: 1 }, { text: "الكمية", rowIndex: 0, columnIndex: 2 }, { text: "السعر", rowIndex: 0, columnIndex: 3 },
    { text: "001", rowIndex: 1, columnIndex: 0 }, { text: "منتج أ", rowIndex: 1, columnIndex: 1 }, { text: "١٠", rowIndex: 1, columnIndex: 2 }, { text: "500 DA", rowIndex: 1, columnIndex: 3 },
    { text: "002", rowIndex: 2, columnIndex: 0 }, { text: "Produit B", rowIndex: 2, columnIndex: 1 }, { text: "5", rowIndex: 2, columnIndex: 2 }, { text: "28/08/2026", rowIndex: 2, columnIndex: 3 },
  ] }], pages: [{ pageNumber: 1, language: "mixed", sourceKind: "text-pdf", text: "", blocks: [] }],
};

describe("Excel conversion", () => {
  it("classifies Arabic digits, dates, percentages, and preserves leading zeros", () => {
    expect(classifyExcelCell("00125")).toEqual({ value: "00125", type: "text" });
    expect(classifyExcelCell("١٢٥٠")).toEqual({ value: 1250, type: "number" });
    expect(classifyExcelCell("15%")).toEqual({ value: 0.15, type: "percentage" });
    expect(classifyExcelCell("28/08/2026").type).toBe("date");
    expect(classifyExcelCell("500 DA").type).toBe("currency");
  });
  it("builds real two-dimensional cell values and separate worksheet mode", () => {
    const plan = buildExcelPlan(fixture, "separate");
    const table = plan.worksheets[0]!.tables[0]!;
    expect(table.values).toEqual([["رقم", "المنتج", "الكمية", "السعر"], ["001", "منتج أ", 10, "500 DA"], ["002", "Produit B", 5, "28/08/2026"]]);
    expect(table.cells.find((cell) => cell.text === "١٠")).toMatchObject({ value: 10, type: "number", rowIndex: 1, columnIndex: 2 });
    expect(table.direction).toBe("rtl");
  });
  it("builds explicit rows and columns with conservative cell formatting", () => {
    const table = buildExcelPlan(fixture, "smart").worksheets[0]!.tables[0]!;
    expect(table.rows).toHaveLength(3); expect(table.columns).toHaveLength(4);
    expect(table.cells.find((cell) => cell.text === "رقم")?.format).toMatchObject({ border: "thin", bold: true, wrapText: true });
  });
  it("creates confirmed merge ranges and does not duplicate covered values", () => {
    const merged: OCRDocument = { ...fixture, tables: [{ ...fixture.tables[0]!, cells: fixture.tables[0]!.cells.map((cell) => cell.text === "رقم" ? { ...cell, merge: { rowSpan: 1, columnSpan: 2, confirmed: true } } : cell) }] };
    const table = buildExcelPlan(merged, "single").worksheets[0]!.tables[0]!;
    expect(table.mergeRanges).toEqual(["A1:B1"]);
    expect(table.values[0]![0]).toBe("رقم");
    expect(table.values[0]![1]).toBeNull();
    const uncertain: OCRDocument = { ...fixture, tables: [{ ...fixture.tables[0]!, cells: fixture.tables[0]!.cells.map((cell) => cell.text === "رقم" ? { ...cell, merge: { rowSpan: 1, columnSpan: 2, confirmed: false } } : cell) }] };
    expect(buildExcelPlan(uncertain, "single").worksheets[0]!.tables[0]!.mergeRanges).toEqual([]);
  });
  it("resolves PDF, image, and smart operation paths without relying on a default", () => {
    expect(resolveExcelOperation("pdf", "invoice.pdf", "application/pdf")).toBe("pdf");
    expect(resolveExcelOperation("image", "invoice.png", "image/png")).toBe("image");
    expect(resolveExcelOperation("smart", "invoice.pdf", "application/pdf")).toBe("smart");
    expect(resolveExcelOperation("pdf", "invoice.png", "image/png")).toBe("smart");
  });
  it("builds the exact Excel transform payload used by the Task Pane", () => {
    const pdf = buildExcelTransformRequest({ name: "arabic.pdf", type: "application/pdf" }, "bytes", "pdf", "separate");
    const image = buildExcelTransformRequest({ name: "mixed.png", type: "image/png" }, "bytes", "image", "single");
    const smart = buildExcelTransformRequest({ name: "mixed.pdf", type: "application/pdf" }, "bytes", "smart", "smart");
    expect(pdf).toMatchObject({ operation: "pdf", mode: "separate" });
    expect(image).toMatchObject({ operation: "image", mode: "single" });
    expect(smart).toMatchObject({ operation: "smart", mode: "smart" });
  });
  it("submits the selected Excel operation to the mutation at runtime", () => {
    const calls: unknown[] = [];
    const file = { name: "arabic.pdf", type: "application/pdf" };
    const request = submitExcelTransform((payload) => { calls.push(payload); }, file, "bytes", "pdf", "separate");
    expect(calls).toHaveLength(1);
    expect(calls[0]).toEqual(request);
    expect(request).toMatchObject({ operation: "pdf", mode: "separate", fileName: "arabic.pdf" });
  });
  it("supports single worksheet mode and safe A1 addresses", () => {
    expect(buildExcelPlan(fixture, "single").worksheets[0]!.name).toBe("Nawa OCR");
    expect(a1Address(0, 0)).toBe("A1");
    expect(a1Address(2, 3)).toBe("D3");
  });
});

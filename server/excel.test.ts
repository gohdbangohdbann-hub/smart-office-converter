import { describe, expect, it } from "vitest";
import { a1Address, buildExcelPlan, classifyExcelCell, inferTablesFromRawCells, mergeInferredTablesAcrossPages, resolveExcelOperation } from "@shared/excel";
import type { OCRDocument } from "@shared/ocr";
import { buildExcelTransformRequest, submitExcelTransform } from "@shared/excel-request";
import { measureExcelFixture } from "@shared/excel-metrics";

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
  it("measures table, row, column, numeric, currency, date, and RTL fixture quality", () => {
    const metrics = measureExcelFixture(buildExcelPlan(fixture, "smart"));
    expect(metrics.cellAccuracy).toBe(1); expect(metrics.rowAccuracy).toBe(1); expect(metrics.columnAccuracy).toBe(1); expect(metrics.tableStructureAccuracy).toBe(1); expect(metrics.numericAccuracy).toBe(1); expect(metrics.currencyAccuracy).toBe(1); expect(metrics.dateAccuracy).toBe(1); expect(metrics.rtlAccuracy).toBeGreaterThan(0.5);
  });
  it("keeps multiple OCR tables on separate worksheets in smart mode", () => {
    const second = { ...fixture.tables[0]!, cells: fixture.tables[0]!.cells.map((cell) => ({ ...cell, rowIndex: cell.rowIndex + 3 })) };
    const plan = buildExcelPlan({ ...fixture, pageCount: 2, tables: [fixture.tables[0]!, second] }, "smart");
    expect(plan.worksheets).toHaveLength(2); expect(plan.worksheets.map((sheet) => sheet.tables)).toHaveLength(2);
  });
  it("infers borderless table rows and columns from raw OCR coordinates", () => {
    const raw = [
      { text: "رقم", pageNumber: 1, polygon: { x: 0.1, y: 0.1, width: 0.1, height: 0.04 } },
      { text: "القيمة", pageNumber: 1, polygon: { x: 0.5, y: 0.1, width: 0.14, height: 0.04 } },
      { text: "001", pageNumber: 1, polygon: { x: 0.1, y: 0.2, width: 0.1, height: 0.04 } },
      { text: "500 DA", pageNumber: 1, polygon: { x: 0.5, y: 0.2, width: 0.14, height: 0.04 } },
      { text: "صفحة 2", pageNumber: 2, polygon: { x: 0.1, y: 0.1, width: 0.2, height: 0.04 } },
    ];
    const inferred = inferTablesFromRawCells(raw);
    expect(inferred).toHaveLength(2); expect(inferred[0]).toMatchObject({ pageStart: 1, pageEnd: 1, rowCount: 2, columnCount: 2 }); expect(inferred[0]!.cells.map((cell) => [cell.rowIndex, cell.columnIndex])).toEqual([[0, 0], [0, 1], [1, 0], [1, 1]]);
    expect(inferred[0]!.cells.filter((cell) => cell.isHeader)).toHaveLength(2);
    const sparse = inferTablesFromRawCells(raw.filter((cell) => cell.text !== "500 DA"));
    expect(sparse[0]!.cells.find((cell) => cell.rowIndex === 1 && cell.columnIndex === 1)).toMatchObject({ text: "", isEmpty: true });
    const repeated = [...raw.slice(0, 4), { text: "رقم", pageNumber: 2, polygon: { x: 0.1, y: 0.1, width: 0.1, height: 0.04 } }, { text: "القيمة", pageNumber: 2, polygon: { x: 0.5, y: 0.1, width: 0.14, height: 0.04 } }, { text: "002", pageNumber: 2, polygon: { x: 0.1, y: 0.2, width: 0.1, height: 0.04 } }, { text: "600 DA", pageNumber: 2, polygon: { x: 0.5, y: 0.2, width: 0.14, height: 0.04 } }];
    const combined = mergeInferredTablesAcrossPages(inferTablesFromRawCells(repeated));
    expect(combined).toHaveLength(1); expect(combined[0]).toMatchObject({ pageStart: 1, pageEnd: 2, rowCount: 3 }); expect(combined[0]!.cells.filter((cell) => cell.text === "رقم")).toHaveLength(1);
  });
  it("supports single worksheet mode and safe A1 addresses", () => {
    expect(buildExcelPlan(fixture, "single").worksheets[0]!.name).toBe("Nawa OCR");
    expect(a1Address(0, 0)).toBe("A1");
    expect(a1Address(2, 3)).toBe("D3");
  });
});

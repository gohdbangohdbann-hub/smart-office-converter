import { describe, expect, it } from "vitest";
import { buildExcelPlan } from "@shared/excel";
import { buildWordPlan } from "@shared/word";
import type { OCRDocument } from "@shared/ocr";

const document: OCRDocument = {
  id: "preview-fixture",
  fileName: "فاتورة.pdf",
  mimeType: "application/pdf",
  language: "ar",
  processedAt: new Date(0).toISOString(),
  pageCount: 1,
  provider: "demo-local",
  privacy: { originalRetained: false, temporaryDataDeleted: true },
  pages: [{ pageNumber: 1, sourceKind: "text-pdf", direction: "rtl", blocks: [{ id: "b1", text: "المجموع ١٢٠", confidence: 0.98, direction: "rtl", kind: "paragraph", lines: [], polygon: [] }, { id: "b2", text: "الوصف المبلغ", confidence: 0.98, direction: "rtl", kind: "table", lines: [], polygon: [], table: { id: "t1", rowCount: 2, columnCount: 2, cells: [{ text: "الوصف", rowIndex: 0, columnIndex: 0, type: "text", language: "ar", confidence: 0.98 }, { text: "المبلغ", rowIndex: 0, columnIndex: 1, type: "text", language: "ar", confidence: 0.98 }, { text: "خدمة", rowIndex: 1, columnIndex: 0, type: "text", language: "ar", confidence: 0.98 }, { text: "١٢٠", rowIndex: 1, columnIndex: 1, type: "currency", language: "ar", confidence: 0.98 }] } }], text: "المجموع ١٢٠" }],
  tables: [{ id: "t1", rowCount: 2, columnCount: 2, cells: [{ text: "الوصف", rowIndex: 0, columnIndex: 0, type: "text", language: "ar", confidence: 0.98 }, { text: "المبلغ", rowIndex: 0, columnIndex: 1, type: "text", language: "ar", confidence: 0.98 }, { text: "خدمة", rowIndex: 1, columnIndex: 0, type: "text", language: "ar", confidence: 0.98 }, { text: "١٢٠", rowIndex: 1, columnIndex: 1, type: "currency", language: "ar", confidence: 0.98 }] }],
};

describe("output preview contract", () => {
  it("uses the Word plan and preserves RTL text before download", () => {
    const plan = buildWordPlan(document);
    expect(plan.pages[0]?.paragraphs[0]?.text).toContain("المجموع");
    expect(plan.pages[0]?.paragraphs[0]?.direction).toBe("rtl");
    expect(plan.pages[0]?.tables[0]?.rows[1]?.[1]).toBe("١٢٠");
  });

  it("uses the Excel plan and exposes worksheet/cell values before download", () => {
    const plan = buildExcelPlan(document, "single");
    const table = plan.worksheets[0]?.tables[0];
    expect(plan.worksheets).toHaveLength(1);
    expect(table?.values[0]).toEqual(["الوصف", "المبلغ"]);
    expect(table?.values[1]?.[1]).toBe(120);
  });
});

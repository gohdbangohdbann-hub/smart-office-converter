import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { buildWordPlan, conservativeArabicPostProcess, detectParagraphDirection } from "@shared/word";
import { progressFor } from "@shared/word-progress";
import { reduceWordProgress } from "@shared/word-progress-flow";
import type { OCRDocument } from "@shared/ocr";

const documentFixture: OCRDocument = {
  id: "doc-1", fileName: "فاتورة.pdf", mimeType: "application/pdf", pageCount: 2, language: "mixed", provider: "fixture", processedAt: new Date().toISOString(),
  privacy: { originalRetained: false, temporaryDataDeleted: true },
  tables: [{ rowCount: 2, columnCount: 2, cells: [{ text: "الرقم", rowIndex: 0, columnIndex: 0 }, { text: "المبلغ", rowIndex: 0, columnIndex: 1 }, { text: "001", rowIndex: 1, columnIndex: 0 }, { text: "5,000 DA", rowIndex: 1, columnIndex: 1 }] }],
  pages: [
    { pageNumber: 1, language: "mixed", sourceKind: "text-pdf", text: "عنوان الفاتورة", confidence: 0.98, blocks: [{ id: "b1", type: "heading", text: "عنوان الفاتورة", language: "ar", confidence: 0.98, lines: [] }] },
    { pageNumber: 2, language: "mixed", sourceKind: "scanned-pdf", text: "فاتورة Invoice", confidence: 0.7, blocks: [{ id: "b2", type: "paragraph", text: "فاتورة   Invoice", language: "mixed", confidence: 0.7, lines: [] }, { id: "b3", type: "table", text: "الرقم المبلغ", language: "ar", confidence: 0.88, lines: [], table: { ...({ rowCount: 2, columnCount: 2, cells: [{ text: "الرقم", rowIndex: 0, columnIndex: 0 }, { text: "المبلغ", rowIndex: 0, columnIndex: 1 }, { text: "001", rowIndex: 1, columnIndex: 0 }, { text: "5,000 DA", rowIndex: 1, columnIndex: 1 }] }) } }] },
  ],
};

describe("Word transformation", () => {
  it("detects paragraph direction and preserves low-confidence text", () => {
    expect(detectParagraphDirection("رقم الفاتورة: 1250", "ar")).toBe("rtl");
    expect(detectParagraphDirection("Date: 28/08/2026", "en")).toBe("ltr");
    expect(conservativeArabicPostProcess("كلمة   غير مؤكدة", 0.4)).toBe("كلمة   غير مؤكدة");
  });
  it("models explicit progress stages and page progress", () => {
    expect(progressFor("analyzing")).toMatchObject({ percent: 15, currentPage: 0 });
    expect(progressFor("ocr", 2, 4)).toMatchObject({ percent: 63, currentPage: 2, totalPages: 4 });
    expect(progressFor("building", 4, 4).message).toContain("Word");
    expect(progressFor("done").percent).toBe(100);
  });
  it("reduces the complete file-to-Word progress sequence", () => {
    let state = reduceWordProgress(progressFor("idle"), { type: "file-selected", totalPages: 4 });
    state = reduceWordProgress(state, { type: "start" });
    expect(state.stage).toBe("analyzing");
    state = reduceWordProgress(state, { type: "stage", stage: "ocr", currentPage: 2 });
    expect(state).toMatchObject({ stage: "ocr", percent: 63, currentPage: 2, totalPages: 4 });
    state = reduceWordProgress(state, { type: "stage", stage: "building", currentPage: 4 });
    expect(state.stage).toBe("building");
    state = reduceWordProgress(state, { type: "success" });
    expect(state).toMatchObject({ stage: "done", percent: 100, currentPage: 4, totalPages: 4 });
    expect(reduceWordProgress(state, { type: "failure" }).stage).toBe("error");
  });
  it("wires the explicit progress model into the Word Task Pane", () => {
    const source = readFileSync(new URL("../client/src/pages/Home.tsx", import.meta.url), "utf8");
    expect(source).toContain("progressFor");
    expect(source).toContain("wordProgress.percent");
    expect(source).toContain("wordProgress.currentPage");
    expect(source).toContain("wordProgress.totalPages");
  });
  it("builds ordered Word pages with editable paragraphs and a real table plan", () => {
    const plan = buildWordPlan(documentFixture);
    expect(plan.pages.map((page) => page.pageNumber)).toEqual([1, 2]);
    expect(plan.pages[0]?.paragraphs[0]).toMatchObject({ kind: "heading", direction: "rtl", bold: true });
    expect(plan.pages[1]?.tables[0]?.rows).toEqual([["الرقم", "المبلغ"], ["001", "5,000 DA"]]);
    expect(plan.pages[1]?.paragraphs[0]?.direction).toBe("rtl");
  });
});

import { describe, expect, it } from "vitest";
import { applyReviewAction, verifyDocument } from "@shared/verification";
import type { OCRDocument } from "@shared/ocr";

function fixture(): OCRDocument {
  return {
    id: "verification-fixture",
    fileName: "arabic-mixed-review.pdf",
    mimeType: "application/pdf",
    pageCount: 1,
    language: "mixed",
    provider: "fixture-provider",
    processedAt: "2026-08-28T00:00:00.000Z",
    privacy: { originalRetained: false, temporaryDataDeleted: true },
    pages: [{ pageNumber: 1, language: "mixed", sourceKind: "mixed-pdf", confidence: 0.86, text: "رقم الفاتورة: 15O00 Date: 28/08/2026", blocks: [{ id: "b1", type: "paragraph", language: "mixed", confidence: 0.72, text: "رقم الفاتورة: 15O00 Date: 28/08/2026", lines: [{ text: "رقم الفاتورة: 15O00 Date: 28/08/2026", confidence: 0.72, words: [{ text: "15O00", confidence: 0.72 }, { text: "رقم", confidence: 0.72 }] }] }] }],
    tables: [{ rowCount: 2, columnCount: 3, pageStart: 1, pageEnd: 1, confidence: 0.91, cells: [
      { text: "الكمية", rowIndex: 0, columnIndex: 0, confidence: 0.98, isHeader: true },
      { text: "السعر", rowIndex: 0, columnIndex: 1, confidence: 0.98, isHeader: true },
      { text: "المجموع", rowIndex: 0, columnIndex: 2, confidence: 0.98, isHeader: true },
      { text: "10", rowIndex: 1, columnIndex: 0, confidence: 0.96 },
      { text: "1500", rowIndex: 1, columnIndex: 1, confidence: 0.96 },
      { text: "15000", rowIndex: 1, columnIndex: 2, confidence: 0.96 },
    ] }],
  };
}

describe("Verification Engine", () => {
  it("detects low confidence, number ambiguity, and mixed RTL content without changing text", () => {
    const report = verifyDocument(fixture());
    expect(report.summary.totalWords).toBe(2);
    expect(report.summary.totalCells).toBe(6);
    expect(report.automaticChanges).toBe(0);
    expect(report.issues.some((item) => item.kind === "low-confidence")).toBe(true);
    const numberIssue = report.issues.find((item) => item.kind === "number");
    expect(numberIssue?.suggestion).toBe("15000");
    expect(report.issues.some((item) => item.kind === "rtl")).toBe(true);
  });

  it("detects arithmetic conflicts while preserving the extracted value", () => {
    const document = fixture();
    document.tables[0] = { rowCount: 2, columnCount: 3, pageStart: 1, pageEnd: 1, confidence: 0.91, cells: [
      { text: "quantity", rowIndex: 0, columnIndex: 0, confidence: 0.98, isHeader: true },
      { text: "price", rowIndex: 0, columnIndex: 1, confidence: 0.98, isHeader: true },
      { text: "total", rowIndex: 0, columnIndex: 2, confidence: 0.98, isHeader: true },
      { text: "10", rowIndex: 1, columnIndex: 0, confidence: 0.96 },
      { text: "1500", rowIndex: 1, columnIndex: 1, confidence: 0.96 },
      { text: "15500", rowIndex: 1, columnIndex: 2, confidence: 0.96 },
    ] };
    const report = verifyDocument(document);
    const arithmetic = report.issues.find((item) => item.kind === "arithmetic");
    expect(arithmetic?.level).toBe("high-risk");
    expect(arithmetic?.extractedText).toBe("15500");
    expect(arithmetic?.suggestion).toBe("15000");
  });

  it("records a user correction without changing the source document", () => {
    const report = verifyDocument(fixture());
    const issue = report.issues.find((item) => item.kind === "number" && item.suggestion === "15000");
    expect(issue).toBeDefined();
    const updated = applyReviewAction(report, { issueId: issue!.id, method: "user-correction", before: issue!.extractedText, after: issue!.suggestion ?? issue!.extractedText, at: "2026-08-28T00:01:00.000Z" });
    expect(updated.issues.find((item) => item.id === issue!.id)?.reviewed).toBe(true);
    expect(fixture().tables[0]!.cells[5]!.text).toBe("15000");
  });
});

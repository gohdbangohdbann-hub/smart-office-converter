import { describe, expect, it } from "vitest";
import { inspectFile, processOCR } from "./ocr";

function pdfWithPages(count: number): string { return Buffer.from(`%PDF-1.7\n${Array.from({ length: count }, () => "/Type /Page").join(" ")}\n`).toString("base64"); }

describe("Phase 5 local engineering smoke benchmark", () => {
  it("measures local inspection and offline processing without accuracy claims", async () => {
    const rows: Array<{ pages: number; inspectMs: number; processMs: number }> = [];
    for (const pages of [1, 10, 50]) {
      const bytesBase64 = pdfWithPages(pages);
      const inspectStart = performance.now();
      const metadata = inspectFile({ fileName: `synthetic-${pages}.pdf`, mimeType: "application/pdf", bytesBase64 });
      const inspectMs = Number((performance.now() - inspectStart).toFixed(2));
      const processStart = performance.now();
      const document = await processOCR({ fileName: `synthetic-${pages}.pdf`, mimeType: "application/pdf", bytesBase64 });
      const processMs = Number((performance.now() - processStart).toFixed(2));
      rows.push({ pages: metadata.pageCount, inspectMs, processMs });
      expect(document.privacy.originalRetained).toBe(false);
    }
    console.info("[Nawa OCR] local smoke benchmark (not OCR accuracy)", rows);
    expect(rows.map(row => row.pages)).toEqual([1, 10, 50]);
  });
});

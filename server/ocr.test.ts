import { describe, expect, it } from "vitest";
import { assertSupportedFile, classifyPdf, createOCRProvider, DemoOCRProvider, processOCR } from "./ocr";
import { detectLanguage, flattenDocumentText, toPreviewModel, type OCRDocument, type OCRProvider } from "@shared/ocr";

describe("Nawa OCR core", () => {
  it("accepts the required PDF and image formats", () => {
    expect(() => assertSupportedFile("document.pdf", "application/pdf")).not.toThrow();
    expect(() => assertSupportedFile("scan.webp", "image/webp")).not.toThrow();
    expect(() => assertSupportedFile("notes.txt", "text/plain")).toThrow("UNSUPPORTED_FILE");
  });

  it("classifies text, scanned, and mixed PDF byte signatures", () => {
    expect(classifyPdf(Buffer.from("/Type /Page BT Arabic ET /Type /Page BT English ET"))).toBe("text-pdf");
    expect(classifyPdf(Buffer.from("/Type /Page /Image /Type /Page"))).toBe("scanned-pdf");
    expect(classifyPdf(Buffer.from("/Type /Page BT Arabic ET /Type /Page /Image"))).toBe("mixed-pdf");
  });

  it("detects Arabic and mixed scripts without forcing a locale", () => {
    expect(detectLanguage("فاتورة عربية")).toBe("ar");
    expect(detectLanguage("فاتورة Invoice 2026")).toBe("mixed");
  });

  it("selects providers through configuration and preserves the complete result contract", async () => {
    const previous = process.env.OCR_PROVIDER;
    process.env.OCR_PROVIDER = "demo";
    const provider = createOCRProvider();
    expect(provider).toBeInstanceOf(DemoOCRProvider);
    const first = await provider.analyze({ fileName: "mixed.png", mimeType: "image/png", bytes: Buffer.from("not-an-image") });
    const secondProvider: OCRProvider = { name: "alternate-test-provider", analyze: async (input) => ({ ...first, provider: "alternate-test-provider", fileName: input.fileName }) };
    const second = await secondProvider.analyze({ fileName: "mixed.png", mimeType: "image/png", bytes: Buffer.from("not-an-image") });
    const withTable = { ...first, tables: [{ rowCount: 1, columnCount: 2, cells: [{ text: "الاسم", rowIndex: 0, columnIndex: 0 }, { text: "Name", rowIndex: 0, columnIndex: 1 }] }] };
    for (const result of [first, second, withTable]) assertOCRDocumentContract(result);
    expect(first.pages.map((page) => page.text)).toEqual(second.pages.map((page) => page.text));
    if (previous === undefined) delete process.env.OCR_PROVIDER; else process.env.OCR_PROVIDER = previous;
  });

  it("returns a unified result and deletes temporary bytes in the offline path", async () => {
    const result = await processOCR({ fileName: "scan.pdf", mimeType: "application/pdf", bytesBase64: Buffer.from("/Type /Page /Image").toString("base64") });
    expect(result.pages).toHaveLength(1);
    expect(result.pages[0]?.sourceKind).toBe("scanned-pdf");
    expect(result.privacy).toEqual({ originalRetained: false, temporaryDataDeleted: true });
    expect(flattenDocumentText(result)).toContain("تجهيز");
    expect(result.pages[0]?.blocks[0]?.confidence).toBeGreaterThan(0);
  });

  it("keeps the preview model stable across provider outputs", async () => {
    const source = await processOCR({ fileName: "arabic.pdf", mimeType: "application/pdf", bytesBase64: Buffer.from("/Type /Page BT Arabic ET").toString("base64") });
    const alternate = { ...source, provider: "alternate-test-provider" };
    expect(toPreviewModel(source)).toEqual(toPreviewModel(alternate));
  });
});

function assertOCRDocumentContract(result: OCRDocument) {
  expect(result.id).toEqual(expect.any(String));
  expect(result.fileName).toEqual(expect.any(String));
  expect(result.mimeType).toEqual(expect.any(String));
  expect(result.language).toEqual(expect.any(String));
  expect(result.processedAt).toEqual(expect.any(String));
  expect(result.pageCount).toBe(result.pages.length);
  expect(result.provider).toEqual(expect.any(String));
  expect(result.privacy).toEqual({ originalRetained: false, temporaryDataDeleted: true });
  expect(result.tables).toEqual(expect.any(Array));
  expect(result.tables.every((table) => Array.isArray(table.cells))).toBe(true);
  for (const cell of result.tables.flatMap((table) => table.cells)) {
    expect(cell.text).toEqual(expect.any(String));
    expect(cell.rowIndex).toEqual(expect.any(Number));
    expect(cell.columnIndex).toEqual(expect.any(Number));
  }
  expect(result.pages[0]?.sourceKind).toEqual(expect.any(String));
  expect(result.pages[0]?.blocks[0]).toEqual(expect.objectContaining({ id: expect.any(String), text: expect.any(String), lines: expect.any(Array) }));
  expect(result.pages[0]?.blocks[0]?.lines[0]?.words).toEqual(expect.any(Array));
}

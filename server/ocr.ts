import { nanoid } from "nanoid";
import { detectLanguage, type OCRBlock, type OCRDocument, type OCRProvider, type OCRRequest } from "@shared/ocr";
import { isImageMime, preprocessImage } from "./preprocess";

const SUPPORTED_TYPES = new Set([
  "application/pdf",
  "image/png",
  "image/jpeg",
  "image/tiff",
  "image/webp",
]);

export function assertSupportedFile(fileName: string, mimeType: string) {
  const extension = fileName.toLowerCase().split(".").pop();
  const supportedExtension = ["pdf", "png", "jpg", "jpeg", "tiff", "webp"].includes(extension ?? "");
  if (!SUPPORTED_TYPES.has(mimeType) && !supportedExtension) {
    throw new Error("UNSUPPORTED_FILE");
  }
}

export function classifyPdf(bytes: Buffer): "text-pdf" | "scanned-pdf" | "mixed-pdf" {
  const source = bytes.toString("latin1");
  const pages = Math.max(1, (source.match(/\/Type\s*\/Page\b/g) ?? []).length);
  const textPages = Math.min(pages, (source.match(/BT[\s\S]{1,2000}?ET/g) ?? []).length);
  if (textPages === 0) return "scanned-pdf";
  if (textPages < pages) return "mixed-pdf";
  return "text-pdf";
}

function makeDemoDocument(input: OCRRequest, bytes: Buffer): OCRDocument {
  const isPdf = input.mimeType === "application/pdf" || input.fileName.toLowerCase().endsWith(".pdf");
  const sourceKind = isPdf ? classifyPdf(bytes) : "image";
  const sample = isPdf && sourceKind === "text-pdf"
    ? "نص مستخرج من ملف PDF نصي. هذا المسار يحافظ على ترتيب الصفحة قبل OCR."
    : "تم تجهيز الملف للتحليل الضوئي. اربط مزود OCR لإرجاع النص الفعلي والثقة والجداول.";
  const language = detectLanguage(sample);
  const line = { text: sample, confidence: 0.74, language, words: sample.split(/\s+/).map((text) => ({ text, confidence: 0.74, language })) };
  const block: OCRBlock = { id: nanoid(8), type: "paragraph", text: sample, language, confidence: 0.74, lines: [line] };
  return {
    id: nanoid(), fileName: input.fileName, mimeType: input.mimeType, pageCount: 1,
    language, pages: [{ pageNumber: 1, language, blocks: [block], text: sample, confidence: 0.74, sourceKind }],
    tables: [], provider: "demo-local", processedAt: new Date().toISOString(),
    privacy: { originalRetained: false, temporaryDataDeleted: true },
  };
}

export class DemoOCRProvider implements OCRProvider {
  readonly name = "demo-local";
  async analyze(input: { fileName: string; mimeType: string; bytes: Buffer }): Promise<OCRDocument> {
    return makeDemoDocument({ fileName: input.fileName, mimeType: input.mimeType, bytesBase64: input.bytes.toString("base64") }, input.bytes);
  }
}

export function createOCRProvider(): OCRProvider {
  const configured = process.env.OCR_PROVIDER?.toLowerCase();
  if (configured === "demo" || (!process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT && !process.env.AZURE_DOCUMENT_INTELLIGENCE_KEY)) return new DemoOCRProvider();
  return new AzureDocumentIntelligenceProvider();
}

export class AzureDocumentIntelligenceProvider implements OCRProvider {
  readonly name = "azure-document-intelligence-layout";
  async analyze(input: { fileName: string; mimeType: string; bytes: Buffer }): Promise<OCRDocument> {
    const endpoint = process.env.AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT;
    const key = process.env.AZURE_DOCUMENT_INTELLIGENCE_KEY;
    if (!endpoint || !key) throw new Error("OCR_PROVIDER_NOT_CONFIGURED");
    const url = `${endpoint.replace(/\/$/, "")}/documentintelligence/documentModels/prebuilt-layout:analyze?api-version=2024-11-30`;
    const start = await fetch(url, { method: "POST", headers: { "Content-Type": input.mimeType, "Ocp-Apim-Subscription-Key": key }, body: input.bytes.buffer.slice(input.bytes.byteOffset, input.bytes.byteOffset + input.bytes.byteLength) as ArrayBuffer });
    if (!start.ok) throw new Error("OCR_SERVICE_ERROR");
    const operation = start.headers.get("operation-location");
    if (!operation) throw new Error("OCR_SERVICE_ERROR");
    const deadline = Date.now() + 150_000;
    while (Date.now() < deadline) {
      await new Promise((resolve) => setTimeout(resolve, 1200));
      const response = await fetch(operation, { headers: { "Ocp-Apim-Subscription-Key": key } });
      if (!response.ok) throw new Error("OCR_SERVICE_ERROR");
      const result = await response.json() as any;
      if (result.status === "succeeded") return normalizeAzureResult(input, result.analyzeResult);
      if (result.status === "failed") throw new Error("OCR_SERVICE_ERROR");
    }
    throw new Error("OCR_TIMEOUT");
  }
}

function normalizeAzureResult(input: { fileName: string; mimeType: string }, result: any): OCRDocument {
  const pages = (result.pages ?? []).map((page: any, index: number) => {
    const lines = (page.lines ?? []).map((line: any) => ({
      text: line.content ?? "", confidence: averageConfidence(line.words), language: detectLanguage(line.content ?? ""),
      words: (line.words ?? []).map((word: any) => ({ text: word.content ?? "", confidence: word.confidence, language: detectLanguage(word.content ?? "") })),
    }));
    const text = lines.map((line: any) => line.text).join("\n");
    const language = detectLanguage(text);
    const block: OCRBlock = { id: nanoid(8), type: "paragraph", text, language, confidence: averageConfidence(lines), lines };
    return { pageNumber: index + 1, language, blocks: [block], text, confidence: averageConfidence(lines), sourceKind: "image" as const };
  });
  const tables = (result.tables ?? []).map((table: any) => ({
    rowCount: table.rowCount ?? 0, columnCount: table.columnCount ?? 0,
    cells: (table.cells ?? []).map((cell: any) => ({ text: cell.content ?? "", rowIndex: cell.rowIndex ?? 0, columnIndex: cell.columnIndex ?? 0, confidence: undefined })),
  }));
  return { id: nanoid(), fileName: input.fileName, mimeType: input.mimeType, pageCount: pages.length, language: detectLanguage(pages.map((p: any) => p.text).join("\n")), pages, tables, provider: "azure-document-intelligence-layout", processedAt: new Date().toISOString(), privacy: { originalRetained: false, temporaryDataDeleted: true } };
}

function averageConfidence(values: any[]): number | undefined {
  const scores = values.flatMap((value) => typeof value === "number" ? [value] : [value?.confidence]).filter((value): value is number => typeof value === "number");
  return scores.length ? Number((scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(3)) : undefined;
}

export function inspectFile(request: Pick<OCRRequest, "fileName" | "mimeType" | "bytesBase64">) {
  assertSupportedFile(request.fileName, request.mimeType);
  const bytes = Buffer.from(request.bytesBase64, "base64");
  if (!bytes.length) throw new Error("EMPTY_FILE");
  const isPdf = request.mimeType === "application/pdf" || request.fileName.toLowerCase().endsWith(".pdf");
  const sourceKind = isPdf ? classifyPdf(bytes) : "image";
  const pageCount = isPdf ? Math.max(1, (bytes.toString("latin1").match(/\/Type\s*\/Page\b/g) ?? []).length) : 1;
  const detectedLanguage = isPdf ? detectLanguage(bytes.toString("utf8")) : "unknown";
  bytes.fill(0);
  return { fileName: request.fileName, mimeType: request.mimeType, pageCount, sourceKind, detectedLanguage };
}

export async function processOCR(request: OCRRequest): Promise<OCRDocument> {
  assertSupportedFile(request.fileName, request.mimeType);
  const bytes = Buffer.from(request.bytesBase64, "base64");
  if (!bytes.length) throw new Error("EMPTY_FILE");
  try {
    const prepared = isImageMime(request.mimeType) ? await preprocessImage(bytes, request.mimeType) : { bytes, steps: [], originalRetained: false as const };
    const provider = createOCRProvider();
    return await provider.analyze({ fileName: request.fileName, mimeType: request.mimeType, bytes: prepared.bytes });
  } catch (error) {
    if (error instanceof Error && error.message === "OCR_PROVIDER_NOT_CONFIGURED") return makeDemoDocument(request, bytes);
    throw error;
  } finally {
    bytes.fill(0);
  }
}

export function userFacingOCRMessage(error: unknown): string {
  const code = error instanceof Error ? error.message : "";
  const messages: Record<string, string> = {
    UNSUPPORTED_FILE: "نوع الملف غير مدعوم. اختر PDF أو PNG أو JPG أو TIFF أو WEBP.",
    EMPTY_FILE: "الملف فارغ أو تعذر قراءته.",
    UNREADABLE_IMAGE: "تعذر قراءة الصورة. استخدم صورة أوضح بصيغة PNG أو JPG.",
    CORRUPT_PDF: "يبدو أن ملف PDF تالف أو محمي. افتحه وأعد حفظه ثم حاول مجددًا.",
    OCR_TIMEOUT: "انتهت مهلة المعالجة. جرّب ملفًا أصغر أو أعد المحاولة.",
    OCR_SERVICE_ERROR: "تعذر إكمال خدمة OCR الآن. تحقق من الاتصال وحاول لاحقًا.",
    OCR_NETWORK_ERROR: "تعذر الاتصال بخدمة OCR. تحقق من الشبكة ثم أعد المحاولة.",
  };
  return messages[code] ?? "حدث خطأ أثناء معالجة الملف. تحقق من الملف ثم أعد المحاولة.";
}

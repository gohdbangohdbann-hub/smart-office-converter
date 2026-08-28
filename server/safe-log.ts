export type OCRLogStatus = "started" | "completed" | "failed";

export function logOCR(event: { mimeType: string; sizeBytes: number; pageCount?: number; durationMs?: number; status: OCRLogStatus; errorCode?: string }): void {
  const safe = { scope: "ocr", ...event, at: new Date().toISOString() };
  if (event.status === "failed") console.warn("[Nawa OCR] operation", safe);
  else console.info("[Nawa OCR] operation", safe);
}

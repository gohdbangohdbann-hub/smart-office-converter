import sharp from "sharp";

export type PreprocessResult = { bytes: Buffer; steps: string[]; originalRetained: false };

export async function preprocessImage(input: Buffer, mimeType: string): Promise<PreprocessResult> {
  if (!input.length) throw new Error("UNREADABLE_IMAGE");
  try {
    const source = sharp(input, { failOn: "error" });
    const metadata = await source.metadata();
    if (!metadata.width || !metadata.height) throw new Error("UNREADABLE_IMAGE");
    const steps = ["auto-rotate", "trim-background", "normalize-contrast", "sharpen", "resize-for-ocr"];
    const base = () => sharp(input).rotate().normalize().sharpen({ sigma: 1.1 }).resize({ width: Math.min(Math.max(metadata.width as number, 1200), 3600), withoutEnlargement: false });
    let pipeline = base().trim({ background: { r: 255, g: 255, b: 255, alpha: 1 }, threshold: 12 });
    let bytes: Buffer;
    try {
      bytes = await pipeline.toFormat(mimeType === "image/png" ? "png" : "jpeg", { quality: 92 }).toBuffer();
    } catch {
      bytes = await base().toFormat(mimeType === "image/png" ? "png" : "jpeg", { quality: 92 }).toBuffer();
    }
    return { bytes, steps, originalRetained: false };
  } catch {
    throw new Error("UNREADABLE_IMAGE");
  }
}

export function isImageMime(mimeType: string) { return mimeType.startsWith("image/"); }
export type PreprocessPdfResult = { bytes: Buffer; steps: string[]; originalRetained: false };
export function preprocessPdf(input: Buffer): PreprocessPdfResult {
  if (!input.length || input.subarray(0, 5).toString("ascii") !== "%PDF-") throw new Error("CORRUPT_PDF");
  return { bytes: Buffer.from(input), steps: ["validate-pdf-signature", "preserve-pdf-for-layout-provider"], originalRetained: false };
}

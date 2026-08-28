import sharp from "sharp";

export type PreprocessResult = {
  bytes: Buffer;
  steps: string[];
  originalRetained: false;
};

export async function preprocessImage(input: Buffer, mimeType: string): Promise<PreprocessResult> {
  if (!input.length) throw new Error("UNREADABLE_IMAGE");
  try {
    const image = sharp(input, { failOn: "error" });
    const metadata = await image.metadata();
    if (!metadata.width || !metadata.height) throw new Error("UNREADABLE_IMAGE");
    const steps = ["auto-rotate", "trim-background", "normalize-contrast", "sharpen", "resize-for-ocr"];
    const bytes = await image.rotate().trim({ background: { r: 255, g: 255, b: 255, alpha: 1 }, threshold: 12 }).normalize().sharpen({ sigma: 1.1 }).resize({ width: Math.min(Math.max(metadata.width, 1200), 3600), withoutEnlargement: false }).toFormat(mimeType === "image/png" ? "png" : "jpeg", { quality: 92 }).toBuffer();
    return { bytes, steps, originalRetained: false };
  } catch {
    throw new Error("UNREADABLE_IMAGE");
  }
}

export function isImageMime(mimeType: string) { return mimeType.startsWith("image/"); }

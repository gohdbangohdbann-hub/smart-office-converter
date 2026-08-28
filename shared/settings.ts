export type OcrLanguage = "auto" | "ar" | "fr" | "en" | "ar-fr" | "ar-en" | "ar-fr-en";
export type NumberDisplay = "arabic" | "western";
export type WordQuality = "highest" | "balanced" | "fast";
export type DocumentDirection = "auto" | "rtl" | "ltr";
export type ExcelTableMode = "separate" | "single" | "smart";
export type ExcelNumericMode = "auto" | "text";

export interface NawaSettings {
  uiLocale: "ar" | "fr" | "en";
  ocrLanguage: OcrLanguage;
  numberDisplay: NumberDisplay;
  wordQuality: WordQuality;
  documentDirection: DocumentDirection;
  fallbackFont: string;
  excelTableMode: ExcelTableMode;
  excelDirection: DocumentDirection;
  excelNumericMode: ExcelNumericMode;
}

export const DEFAULT_SETTINGS: NawaSettings = {
  uiLocale: "ar",
  ocrLanguage: "auto",
  numberDisplay: "western",
  wordQuality: "balanced",
  documentDirection: "auto",
  fallbackFont: "Aptos",
  excelTableMode: "smart",
  excelDirection: "auto",
  excelNumericMode: "auto",
};

export const SETTINGS_STORAGE_KEY = "nawa-ocr-settings-v1";

export function parseSettings(value: string | null): NawaSettings {
  if (!value) return DEFAULT_SETTINGS;
  try {
    const parsed = JSON.parse(value) as Partial<NawaSettings>;
    return { ...DEFAULT_SETTINGS, ...parsed };
  } catch {
    return DEFAULT_SETTINGS;
  }
}

export function formatDigits(value: string, display: NumberDisplay): string {
  if (display === "western") return value.replace(/[٠-٩]/g, d => String("٠١٢٣٤٥٦٧٨٩".indexOf(d)));
  return value.replace(/[0-9]/g, d => "٠١٢٣٤٥٦٧٨٩"[Number(d)] ?? d);
}

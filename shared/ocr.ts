export type OCRLanguage = "ar" | "fr" | "en" | "mixed" | "unknown";
export type OCRBlockType = "paragraph" | "heading" | "list" | "table" | "image" | "separator" | "unknown";

export interface BoundingBox {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface OCRWord {
  text: string;
  confidence?: number;
  language?: OCRLanguage;
  polygon?: BoundingBox;
}

export interface OCRLine {
  text: string;
  confidence?: number;
  language?: OCRLanguage;
  words: OCRWord[];
  polygon?: BoundingBox;
}

export interface OCRCell {
  text: string;
  rowIndex: number;
  columnIndex: number;
  confidence?: number;
  polygon?: BoundingBox;
  merge?: { rowSpan: number; columnSpan: number; confirmed?: boolean };
}

export interface OCRTable {
  rowCount: number;
  columnCount: number;
  cells: OCRCell[];
  confidence?: number;
}

export interface OCRBlock {
  id: string;
  type: OCRBlockType;
  text: string;
  language: OCRLanguage;
  confidence?: number;
  polygon?: BoundingBox;
  lines: OCRLine[];
  table?: OCRTable;
  assetRef?: string;
  listLevel?: number;
  isHeader?: boolean;
  isFooter?: boolean;
}

export interface OCRPage {
  pageNumber: number;
  width?: number;
  height?: number;
  language: OCRLanguage;
  blocks: OCRBlock[];
  text: string;
  confidence?: number;
  sourceKind: "text-pdf" | "scanned-pdf" | "image" | "mixed-pdf";
  header?: OCRBlock[];
  footer?: OCRBlock[];
}

export interface OCRDocument {
  id: string;
  fileName: string;
  mimeType: string;
  pageCount: number;
  language: OCRLanguage;
  pages: OCRPage[];
  tables: OCRTable[];
  provider: string;
  processedAt: string;
  privacy: {
    originalRetained: false;
    temporaryDataDeleted: boolean;
  };
}

export interface OCRRequest {
  fileName: string;
  mimeType: string;
  bytesBase64: string;
}

export interface OCRProvider {
  readonly name: string;
  analyze(input: { fileName: string; mimeType: string; bytes: Buffer }): Promise<OCRDocument>;
}

export function detectLanguage(text: string): OCRLanguage {
  const arabic = (text.match(/[\u0600-\u06ff]/g) ?? []).length;
  const latin = (text.match(/[A-Za-zÀ-ÿ]/g) ?? []).length;
  if (arabic > 0 && latin > 0) return "mixed";
  if (arabic > 0) return "ar";
  if (latin > 0) return "en";
  return "unknown";
}

export function flattenDocumentText(document: OCRDocument): string {
  return document.pages.map((page) => page.text).filter(Boolean).join("\n\n");
}

export interface OCRPreviewModel {
  text: string;
  language: OCRLanguage;
  pageCount: number;
  confidence?: number;
  tables: OCRTable[];
}

export function toPreviewModel(document: OCRDocument): OCRPreviewModel {
  return { text: flattenDocumentText(document), language: document.language, pageCount: document.pageCount, confidence: document.pages[0]?.confidence, tables: document.tables };
}

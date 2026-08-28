import type { OCRBlock, OCRDocument, OCRLanguage, OCRTable } from "./ocr";

export type WordDirection = "rtl" | "ltr";
export type WordBlockKind = "heading" | "paragraph" | "list" | "table" | "image";

export interface WordParagraphPlan {
  kind: Exclude<WordBlockKind, "table" | "image">;
  text: string;
  direction: WordDirection;
  language: OCRLanguage;
  confidence?: number;
  level?: number;
  bold?: boolean;
}

export interface WordTablePlan { direction: WordDirection; rows: string[][]; confidence?: number; }
export interface WordPagePlan { pageNumber: number; paragraphs: WordParagraphPlan[]; tables: WordTablePlan[]; }
export interface WordDocumentPlan { fileName: string; pageCount: number; pages: WordPagePlan[]; warnings: string[]; }

export function detectParagraphDirection(text: string, language: OCRLanguage): WordDirection {
  if (language === "ar") return "rtl";
  if (language === "fr" || language === "en") return "ltr";
  const firstStrong = text.match(/[\u0600-\u06ff]|[A-Za-zÀ-ÿ]/);
  return firstStrong && /[\u0600-\u06ff]/.test(firstStrong[0]) ? "rtl" : "ltr";
}

export function conservativeArabicPostProcess(text: string, confidence = 1) {
  if (confidence < 0.65) return text;
  return text.replace(/[ \t]+/g, " ").replace(/\s+([،؛؟,:.!؟])/g, "$1").trim();
}

function tableRows(table: OCRTable): string[][] {
  return Array.from({ length: table.rowCount }, (_, row) => Array.from({ length: table.columnCount }, (_, column) => table.cells.find((cell) => cell.rowIndex === row && cell.columnIndex === column)?.text ?? ""));
}

function blockPlan(block: OCRBlock): WordParagraphPlan | null {
  if (block.type === "table" || block.type === "image") return null;
  const text = conservativeArabicPostProcess(block.text, block.confidence ?? 1);
  if (!text) return null;
  const heading = block.type === "heading" || /^((الفصل|الباب|chapter|section)\b|\d+[.)]\s)/i.test(text);
  const list = /^([-•*]|\d+[.)])\s/.test(text);
  return { kind: heading ? "heading" : list ? "list" : "paragraph", text, direction: detectParagraphDirection(text, block.language), language: block.language, confidence: block.confidence, level: heading ? 1 : undefined, bold: heading };
}

export function buildWordPlan(document: OCRDocument): WordDocumentPlan {
  return { fileName: document.fileName, pageCount: document.pageCount, pages: document.pages.map((page) => ({ pageNumber: page.pageNumber, paragraphs: page.blocks.map(blockPlan).filter((block): block is WordParagraphPlan => Boolean(block)), tables: page.blocks.filter((block) => block.table).map((block) => ({ direction: detectParagraphDirection(block.text, block.language), rows: tableRows(block.table as OCRTable), confidence: block.confidence })) })), warnings: document.pages.filter((page) => (page.confidence ?? 1) < 0.65).map((page) => `الصفحة ${page.pageNumber} تحتاج مراجعة بسبب انخفاض الثقة`) };
}

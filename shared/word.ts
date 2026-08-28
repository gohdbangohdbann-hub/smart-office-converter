import type { OCRBlock, OCRDocument, OCRLanguage, OCRTable } from "./ocr";

export type WordDirection = "rtl" | "ltr";
export type WordBlockKind = "heading" | "paragraph" | "list" | "table" | "image" | "separator";
export interface WordParagraphPlan { kind: Exclude<WordBlockKind, "table" | "image" | "separator">; text: string; direction: WordDirection; language: OCRLanguage; confidence?: number; level?: number; bold?: boolean; }
export interface WordTablePlan { direction: WordDirection; rows: string[][]; confidence?: number; }
export interface WordAssetPlan { kind: "image" | "separator"; assetRef?: string; }
export interface WordPagePlan { pageNumber: number; paragraphs: WordParagraphPlan[]; tables: WordTablePlan[]; assets: WordAssetPlan[]; header: WordParagraphPlan[]; footer: WordParagraphPlan[]; }
export interface WordDocumentPlan { fileName: string; pageCount: number; pages: WordPagePlan[]; warnings: string[]; }

export function detectParagraphDirection(text: string, language: OCRLanguage): WordDirection {
  if (language === "ar") return "rtl";
  if (language === "fr" || language === "en") return "ltr";
  const firstStrong = text.match(/[\u0600-\u06ff]|[A-Za-zÀ-ÿ]/);
  return firstStrong && /[\u0600-\u06ff]/.test(firstStrong[0]) ? "rtl" : "ltr";
}
export function conservativeArabicPostProcess(text: string, confidence = 1) { return confidence < 0.65 ? text : text.replace(/[ \t]+/g, " ").replace(/\s+([،؛؟,:.!؟])/g, "$1").trim(); }
function tableRows(table: OCRTable): string[][] { return Array.from({ length: table.rowCount }, (_, row) => Array.from({ length: table.columnCount }, (_, column) => table.cells.find((cell) => cell.rowIndex === row && cell.columnIndex === column)?.text ?? "")); }
function blockPlan(block: OCRBlock): WordParagraphPlan | null {
  if (["table", "image", "separator"].includes(block.type)) return null;
  const text = conservativeArabicPostProcess(block.text, block.confidence ?? 1); if (!text) return null;
  const heading = block.type === "heading" || /^((الفصل|الباب|chapter|section)\b|\d+[.)]\s)/i.test(text);
  const list = block.type === "list" || /^([-•*]|\d+[.)])\s/.test(text);
  return { kind: heading ? "heading" : list ? "list" : "paragraph", text, direction: detectParagraphDirection(text, block.language), language: block.language, confidence: block.confidence, level: block.listLevel ?? (heading ? 1 : undefined), bold: heading };
}
function readingOrder(blocks: OCRBlock[], language: OCRLanguage): OCRBlock[] {
  const positioned = blocks.filter((block) => block.polygon).map((block) => ({ block, x: block.polygon!.x + block.polygon!.width / 2, y: block.polygon!.y }));
  const unpositioned = blocks.filter((block) => !block.polygon);
  if (positioned.length < 2) return blocks;
  const columns: Array<typeof positioned> = [];
  for (const item of positioned.sort((a, b) => a.x - b.x || a.y - b.y)) {
    const column = columns.find((candidate) => Math.abs(candidate[0]!.x - item.x) <= 0.2);
    if (column) column.push(item); else columns.push([item]);
  }
  const rtl = language === "ar" || blocks.some((block) => block.language === "ar" || /[\u0600-\u06ff]/.test(block.text));
  const ordered = columns.sort((a, b) => rtl ? b[0]!.x - a[0]!.x : a[0]!.x - b[0]!.x).flatMap((column) => column.sort((a, b) => a.y - b.y).map((item) => item.block));
  return [...ordered, ...unpositioned];
}
function pageParagraphs(blocks: OCRBlock[], language: OCRLanguage) { return readingOrder(blocks, language).map(blockPlan).filter((block): block is WordParagraphPlan => Boolean(block)); }
export function buildWordPlan(document: OCRDocument): WordDocumentPlan {
  return { fileName: document.fileName, pageCount: document.pageCount, pages: document.pages.map((page) => ({ pageNumber: page.pageNumber, paragraphs: pageParagraphs(page.blocks, page.language), tables: page.blocks.filter((block) => block.table).map((block) => ({ direction: detectParagraphDirection(block.text, block.language), rows: tableRows(block.table as OCRTable), confidence: block.confidence })), assets: page.blocks.filter((block) => block.type === "image" || block.type === "separator").map((block) => ({ kind: block.type as "image" | "separator", assetRef: block.assetRef })), header: page.header ? pageParagraphs(page.header, page.language) : [], footer: page.footer ? pageParagraphs(page.footer, page.language) : [] })), warnings: document.pages.filter((page) => (page.confidence ?? 1) < 0.65).map((page) => `الصفحة ${page.pageNumber} تحتاج مراجعة بسبب انخفاض الثقة`) };
}

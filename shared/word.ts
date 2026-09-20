import type { OCRBlock, OCRDocument, OCRLanguage, OCRTable } from "./ocr";

export type WordDirection = "rtl" | "ltr";
export type WordBlockKind = "heading" | "paragraph" | "list" | "table" | "image" | "separator";
export type WordParagraphStyle = "title" | "heading1" | "heading2" | "body" | "list" | "caption" | "header" | "footer";

export interface WordParagraphPlan {
  kind: Exclude<WordBlockKind, "table" | "image" | "separator">;
  style: WordParagraphStyle;
  text: string;
  direction: WordDirection;
  language: OCRLanguage;
  confidence?: number;
  level?: number;
  bold?: boolean;
  italic?: boolean;
  keepWithNext?: boolean;
  indentLevel?: number;
  spacingBefore?: number;
  spacingAfter?: number;
}
export interface WordTablePlan { direction: WordDirection; rows: string[][]; headerRows: number; confidence?: number; }
export interface WordAssetPlan { kind: "image" | "separator"; assetRef?: string; }
export interface WordPagePlan { pageNumber: number; sectionIndex: number; paragraphs: WordParagraphPlan[]; tables: WordTablePlan[]; assets: WordAssetPlan[]; header: WordParagraphPlan[]; footer: WordParagraphPlan[]; }
export interface WordSectionPlan { sectionIndex: number; pageStart: number; pageEnd: number; direction: WordDirection; header: WordParagraphPlan[]; footer: WordParagraphPlan[]; }
export interface WordStyleProfile { defaultDirection: WordDirection; title?: string; headingCount: number; listCount: number; tableCount: number; hasHeaders: boolean; hasFooters: boolean; }
export interface WordDocumentPlan { fileName: string; pageCount: number; pages: WordPagePlan[]; sections: WordSectionPlan[]; styles: WordStyleProfile; warnings: string[]; }

export function detectParagraphDirection(text: string, language: OCRLanguage): WordDirection {
  if (language === "ar") return "rtl";
  if (language === "fr" || language === "en") return "ltr";
  const firstStrong = text.match(/[\u0600-\u06ff]|[A-Za-zÀ-ÿ]/);
  return firstStrong && /[\u0600-\u06ff]/.test(firstStrong[0]) ? "rtl" : "ltr";
}
export function conservativeArabicPostProcess(text: string, confidence = 1) { return confidence < 0.65 ? text : text.replace(/[ \t]+/g, " ").replace(/\s+([،؛؟,:.!؟])/g, "$1").trim(); }
function tableRows(table: OCRTable): string[][] { return Array.from({ length: table.rowCount }, (_, row) => Array.from({ length: table.columnCount }, (_, column) => table.cells.find((cell) => cell.rowIndex === row && cell.columnIndex === column)?.text ?? "")); }
function tableHeaderRows(table: OCRTable): number {
  const explicit = table.cells.filter((cell) => cell.isHeader).map((cell) => cell.rowIndex);
  if (explicit.length) return Math.max(...explicit) + 1;
  const first = table.cells.filter((cell) => cell.rowIndex === 0);
  return first.length > 0 && first.some((cell) => /[A-Za-z\u0600-\u06ff]/.test(cell.text) && !/^[-+]?\d[\d\s.,/%-]*$/.test(cell.text.trim())) ? 1 : 0;
}
function inferHeadingLevel(text: string, block: OCRBlock, firstContentHeading: boolean): 1 | 2 | undefined {
  if (block.type !== "heading" && !/^((الفصل|الباب|chapter|section)\b)/i.test(text) && !/^\d+[.)]\s/.test(text)) return undefined;
  if (firstContentHeading) return 1;
  if (block.listLevel && block.listLevel > 1) return 2;
  return /^\d+[.)]\s/.test(text) ? 2 : 1;
}
function blockPlan(block: OCRBlock, context: { firstContentHeading: boolean; header?: boolean; footer?: boolean }): WordParagraphPlan | null {
  if (["table", "image", "separator"].includes(block.type)) return null;
  const text = conservativeArabicPostProcess(block.text, block.confidence ?? 1); if (!text) return null;
  const headingLevel = inferHeadingLevel(text, block, context.firstContentHeading);
  const list = block.type === "list" || /^([-•*]|\d+[.)])\s/.test(text);
  const style: WordParagraphStyle = context.header ? "header" : context.footer ? "footer" : headingLevel === 1 && context.firstContentHeading ? "title" : headingLevel === 1 ? "heading1" : headingLevel === 2 ? "heading2" : list ? "list" : block.type === "unknown" ? "caption" : "body";
  const kind = style === "list" ? "list" : style === "body" || style === "caption" ? "paragraph" : "heading";
  return { kind, style, text, direction: detectParagraphDirection(text, block.language), language: block.language, confidence: block.confidence, level: headingLevel ?? block.listLevel, bold: ["title", "heading1", "heading2"].includes(style) || block.isHeader === true, italic: style === "caption", keepWithNext: ["title", "heading1", "heading2"].includes(style), indentLevel: style === "list" ? Math.max(0, (block.listLevel ?? 1) - 1) : undefined, spacingBefore: style === "title" ? 0 : style.startsWith("heading") ? 10 : 4, spacingAfter: style.startsWith("heading") || style === "title" ? 5 : 3 };
}
function readingOrder(blocks: OCRBlock[], language: OCRLanguage): OCRBlock[] {
  const positioned = blocks.filter((block) => block.polygon).map((block) => ({ block, x: block.polygon!.x + block.polygon!.width / 2, y: block.polygon!.y }));
  const unpositioned = blocks.filter((block) => !block.polygon); if (positioned.length < 2) return blocks;
  const columns: Array<typeof positioned> = [];
  for (const item of positioned.sort((a, b) => a.x - b.x || a.y - b.y)) { const column = columns.find((candidate) => Math.abs(candidate[0]!.x - item.x) <= 0.2); if (column) column.push(item); else columns.push([item]); }
  const rtl = language === "ar" || blocks.some((block) => block.language === "ar" || /[\u0600-\u06ff]/.test(block.text));
  return [...columns.sort((a, b) => rtl ? b[0]!.x - a[0]!.x : a[0]!.x - b[0]!.x).flatMap((column) => column.sort((a, b) => a.y - b.y).map((item) => item.block)), ...unpositioned];
}
function pageParagraphs(blocks: OCRBlock[], language: OCRLanguage, header = false, footer = false) {
  let firstContentHeading = !header && !footer;
  return readingOrder(blocks, language).map((block) => { const plan = blockPlan(block, { firstContentHeading, header, footer }); if (plan?.style === "title" || plan?.style === "heading1") firstContentHeading = false; return plan; }).filter((block): block is WordParagraphPlan => Boolean(block));
}
function sameParagraphs(a: WordParagraphPlan[], b: WordParagraphPlan[]): boolean { return a.length === b.length && a.every((item, index) => item.text === b[index]?.text && item.direction === b[index]?.direction); }
export function buildWordPlan(document: OCRDocument): WordDocumentPlan {
  const pages = document.pages.map((page, index) => {
    const paragraphs = pageParagraphs(page.blocks, page.language); const header = page.header ? pageParagraphs(page.header, page.language, true) : []; const footer = page.footer ? pageParagraphs(page.footer, page.language, false, true) : [];
    return { pageNumber: page.pageNumber, sectionIndex: index, paragraphs, tables: page.blocks.filter((block) => block.table).map((block) => ({ direction: detectParagraphDirection(block.text, block.language), rows: tableRows(block.table as OCRTable), headerRows: tableHeaderRows(block.table as OCRTable), confidence: block.confidence })), assets: page.blocks.filter((block) => block.type === "image" || block.type === "separator").map((block) => ({ kind: block.type as "image" | "separator", assetRef: block.assetRef })), header, footer };
  });
  const sections: WordSectionPlan[] = [];
  for (const page of pages) { const previous = sections[sections.length - 1]; const direction = page.paragraphs[0]?.direction ?? detectParagraphDirection(document.pages[page.pageNumber - 1]?.text ?? "", document.language); if (previous && sameParagraphs(previous.header, page.header) && sameParagraphs(previous.footer, page.footer) && previous.direction === direction) previous.pageEnd = page.pageNumber; else sections.push({ sectionIndex: sections.length, pageStart: page.pageNumber, pageEnd: page.pageNumber, direction, header: page.header, footer: page.footer }); page.sectionIndex = sections[sections.length - 1]!.sectionIndex; }
  const allParagraphs = pages.flatMap((page) => page.paragraphs);
  return { fileName: document.fileName, pageCount: document.pageCount, pages, sections, styles: { defaultDirection: allParagraphs.find((paragraph) => paragraph.direction)?.direction ?? "ltr", title: allParagraphs.find((paragraph) => paragraph.style === "title")?.text, headingCount: allParagraphs.filter((paragraph) => paragraph.style.startsWith("heading") || paragraph.style === "title").length, listCount: allParagraphs.filter((paragraph) => paragraph.style === "list").length, tableCount: pages.reduce((sum, page) => sum + page.tables.length, 0), hasHeaders: sections.some((section) => section.header.length > 0), hasFooters: sections.some((section) => section.footer.length > 0) }, warnings: document.pages.filter((page) => (page.confidence ?? 1) < 0.65).map((page) => `الصفحة ${page.pageNumber} تحتاج مراجعة بسبب انخفاض الثقة`) };
}

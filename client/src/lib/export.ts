import { AlignmentType, Document, HeadingLevel, Packer, Paragraph, Table, TableCell, TableRow, TextRun } from "docx";
import * as XLSX from "xlsx";
import { buildExcelPlan, type ExcelWorkbookPlan } from "@shared/excel";
import { buildWordPlan, type WordDocumentPlan } from "@shared/word";
import { formatDigits, type NumberDisplay } from "@shared/settings";
import { nextOutputFileName, type BatchOutputMode } from "@shared/batch";
import type { OCRDocument } from "@shared/ocr";

function download(blob: Blob, name: string): void {
  const url = URL.createObjectURL(blob); const anchor = document.createElement("a"); anchor.href = url; anchor.download = name; anchor.click(); window.setTimeout(() => URL.revokeObjectURL(url), 0);
}
function tableFromRows(rows: string[][], headerRows: number, direction: "rtl" | "ltr", numberDisplay: NumberDisplay): Table {
  return new Table({
    rows: rows.map((row, rowIndex) => new TableRow({
      children: row.map((cell) => new TableCell({
        children: [new Paragraph({
          children: [new TextRun({ text: formatDigits(cell, numberDisplay), bold: rowIndex < headerRows })],
          alignment: direction === "rtl" ? AlignmentType.RIGHT : AlignmentType.LEFT,
          bidirectional: direction === "rtl",
        })],
      })),
    })),
  });
}
function paragraphFromPlan(paragraph: WordDocumentPlan["pages"][number]["paragraphs"][number], numberDisplay: NumberDisplay, pageBreakBefore = false): Paragraph {
  const heading = paragraph.style === "title" ? HeadingLevel.TITLE : paragraph.style === "heading1" ? HeadingLevel.HEADING_1 : paragraph.style === "heading2" ? HeadingLevel.HEADING_2 : undefined;
  return new Paragraph({ children: [new TextRun({ text: formatDigits(paragraph.text, numberDisplay), bold: paragraph.bold, italics: paragraph.italic })], alignment: paragraph.direction === "rtl" ? AlignmentType.RIGHT : AlignmentType.LEFT, bidirectional: paragraph.direction === "rtl", heading, pageBreakBefore, keepNext: paragraph.keepWithNext, spacing: { before: paragraph.spacingBefore, after: paragraph.spacingAfter, line: 276 }, indent: paragraph.indentLevel ? { left: paragraph.indentLevel * 360, right: paragraph.direction === "rtl" ? paragraph.indentLevel * 360 : undefined } : undefined });
}
export async function createWordBlob(document: OCRDocument, numberDisplay: NumberDisplay): Promise<Blob> {
  const plan: WordDocumentPlan = buildWordPlan(document); const children: Array<Paragraph | Table> = [];
  for (let pageIndex = 0; pageIndex < plan.pages.length; pageIndex += 1) { const page = plan.pages[pageIndex]!; for (const paragraph of page.header) children.push(paragraphFromPlan(paragraph, numberDisplay, pageIndex > 0)); for (const paragraph of page.paragraphs) children.push(paragraphFromPlan(paragraph, numberDisplay, pageIndex > 0 && page.paragraphs[0] === paragraph)); for (const table of page.tables) children.push(tableFromRows(table.rows, table.headerRows, table.direction, numberDisplay)); for (const paragraph of page.footer) children.push(paragraphFromPlan(paragraph, numberDisplay)); }
  return Packer.toBlob(new Document({ sections: [{ children }] }));
}

export async function downloadWord(document: OCRDocument, numberDisplay: NumberDisplay, existing: string[] = []): Promise<string> {
  const name = nextOutputFileName(document.fileName, "Word", existing); download(await createWordBlob(document, numberDisplay), name); return name;
}
export function combineDocuments(documents: OCRDocument[], fileName: string): OCRDocument {
  let pageNumber = 1;
  return { ...(documents[0] ?? { id: "combined", mimeType: "application/pdf", language: "mixed", tables: [], pages: [], provider: "demo-local", processedAt: new Date().toISOString(), privacy: { originalRetained: false, temporaryDataDeleted: true } }), id: `combined-${Date.now()}`, fileName, pageCount: documents.reduce((sum, item) => sum + item.pageCount, 0), pages: documents.flatMap(item => item.pages.map(page => ({ ...page, pageNumber: pageNumber++ }))), tables: documents.flatMap(item => item.tables) };
}
export async function downloadWordBatch(documents: OCRDocument[], numberDisplay: NumberDisplay, mode: BatchOutputMode, existing: string[] = []): Promise<string[]> {
  if (mode === "separate") { const names: string[] = []; for (const document of documents) names.push(await downloadWord(document, numberDisplay, existing)); return names; }
  const combined = combineDocuments(documents, "Nawa_Batch.docx"); return [await downloadWord(combined, numberDisplay, existing)];
}
export function downloadExcelBatch(documents: OCRDocument[], numberDisplay: NumberDisplay, mode: BatchOutputMode, excelMode: "separate" | "single" | "smart" = "smart", existing: string[] = []): string[] {
  if (mode === "separate") return documents.map(document => downloadExcel(document, numberDisplay, excelMode, existing));
  if (mode === "combined") return [downloadExcel(combineDocuments(documents, "Nawa_Batch.xlsx"), numberDisplay, "single", existing)];
  const workbook = XLSX.utils.book_new(); const names: string[] = [];
  documents.forEach((document, index) => { const plan = buildExcelPlan(document, excelMode); const rows = plan.worksheets.flatMap(worksheet => worksheet.tables.flatMap(table => table.values.map(row => row.map(value => value == null ? "" : typeof value === "number" ? value : formatDigits(String(value), numberDisplay))))); XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), `File ${index + 1}`); });
  const bytes = XLSX.write(workbook, { bookType: "xlsx", type: "array" }); const name = nextOutputFileName("Nawa_Batch", "Excel", existing); download(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), name); names.push(name); return names;
}
export function createExcelBlob(document: OCRDocument, numberDisplay: NumberDisplay, mode: "separate" | "single" | "smart" = "smart"): Blob {
  const plan: ExcelWorkbookPlan = buildExcelPlan(document, mode); const workbook = XLSX.utils.book_new();
  for (const worksheet of plan.worksheets) { const rows = worksheet.tables.flatMap(table => table.values.map(row => row.map(value => value == null ? "" : typeof value === "number" ? value : formatDigits(String(value), numberDisplay)))); const sheet = XLSX.utils.aoa_to_sheet(rows); XLSX.utils.book_append_sheet(workbook, sheet, worksheet.name.slice(0, 31)); }
  const bytes = XLSX.write(workbook, { bookType: "xlsx", type: "array" }); return new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

export function downloadExcel(document: OCRDocument, numberDisplay: NumberDisplay, mode: "separate" | "single" | "smart" = "smart", existing: string[] = []): string {
  const name = nextOutputFileName(document.fileName, "Excel", existing); download(createExcelBlob(document, numberDisplay, mode), name); return name;
}

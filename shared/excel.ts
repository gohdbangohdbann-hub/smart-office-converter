import type { BoundingBox, OCRCell, OCRDocument, OCRLanguage, OCRTable } from "./ocr";

export type ExcelCellType = "text" | "number" | "date" | "percentage" | "currency";
export type ExcelValue = string | number | null;
export interface ExcelCellFormat { horizontalAlignment: "left" | "right"; wrapText: boolean; bold: boolean; numberFormat?: string; border: "thin" | "none"; }
export interface ExcelColumnPlan { index: number; widthHint: number; cells: ExcelCellPlan[]; }
export interface ExcelRowPlan { index: number; heightHint: number; cells: ExcelCellPlan[]; }
export interface ExcelCellPlan { text: string; value: ExcelValue; rowIndex: number; columnIndex: number; type: ExcelCellType; language: OCRLanguage; confidence?: number; polygon?: BoundingBox; merge?: { rowSpan: number; columnSpan: number; confirmed: boolean }; format: ExcelCellFormat; }
export interface ExcelTablePlan { name: string; sheetName: string; rowCount: number; columnCount: number; headers: string[]; rows: ExcelRowPlan[]; columns: ExcelColumnPlan[]; cells: ExcelCellPlan[]; values: ExcelValue[][]; direction: "rtl" | "ltr"; confidence?: number; mergeRanges: string[]; }
export interface ExcelWorksheetPlan { name: string; tables: ExcelTablePlan[]; }
export interface ExcelWorkbookPlan { fileName: string; worksheets: ExcelWorksheetPlan[]; pageCount: number; language: OCRLanguage; warnings: string[]; }
export function mergeInferredTablesAcrossPages(tables: InferredOCRTable[]): InferredOCRTable[] {
  const merged: InferredOCRTable[] = [];
  for (const current of tables) {
    const previous = merged[merged.length - 1];
    const currentHeader = current.cells.filter((cell) => cell.rowIndex === 0).sort((a, b) => a.columnIndex - b.columnIndex).map((cell) => cell.text).join("|");
    const previousHeader = previous?.cells.filter((cell) => cell.rowIndex === 0).sort((a, b) => a.columnIndex - b.columnIndex).map((cell) => cell.text).join("|");
    if (previous && previous.pageEnd + 1 === current.pageStart && previous.columnCount === current.columnCount && currentHeader && currentHeader === previousHeader) {
      const repeatedHeaderRows = current.cells.filter((cell) => cell.rowIndex === 0);
      const body = current.cells.filter((cell) => cell.rowIndex > 0).map((cell) => ({ ...cell, rowIndex: cell.rowIndex - 1 + previous.rowCount }));
      previous.cells = [...previous.cells, ...body]; previous.rowCount += Math.max(0, current.rowCount - 1); previous.pageEnd = current.pageEnd;
      void repeatedHeaderRows;
    } else merged.push({ ...current, cells: [...current.cells] });
  }
  return merged;
}

const ARABIC_DIGITS = "٠١٢٣٤٥٦٧٨٩";
function normalizeDigits(text: string) { return text.replace(/[٠-٩]/g, (digit) => String(ARABIC_DIGITS.indexOf(digit))); }
function direction(language: OCRLanguage, text: string): "rtl" | "ltr" { if (language === "ar") return "rtl"; if (language === "fr" || language === "en") return "ltr"; return /[\u0600-\u06ff]/.test(text) ? "rtl" : "ltr"; }
export function classifyExcelCell(text: string): { value: ExcelValue; type: ExcelCellType } {
  const raw = text.trim(); const normalized = normalizeDigits(raw);
  if (/^0\d+/.test(normalized)) return { value: raw, type: "text" };
  if (/^\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}$/.test(normalized)) return { value: raw, type: "date" };
  if (/^[-+]?\d+(?:[.,]\d+)?%$/.test(normalized)) return { value: Number(normalized.replace(",", ".").replace("%", "")) / 100, type: "percentage" };
  if (/^(?:DZD|DA|EUR|USD|GBP|SAR|AED|MAD|TND)\s*[-+]?\d|[-+]?\d[\d\s.,]*\s*(?:DZD|DA|EUR|GBP|SAR|AED|MAD|TND)$/i.test(normalized)) return { value: raw, type: "currency" };
  if (/^[-+]?\d+(?:[.,]\d+)?$/.test(normalized)) return { value: Number(normalized.replace(",", ".")), type: "number" };
  return { value: raw || null, type: "text" };
}
function excelColumn(index: number) { let result = ""; for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) result = String.fromCharCode(65 + ((n - 1) % 26)) + result; return result; }
export function a1Address(rowIndex: number, columnIndex: number) { return `${excelColumn(columnIndex)}${rowIndex + 1}`; }
function formatFor(cell: OCRCell, language: OCRLanguage, rowIndex: number): ExcelCellFormat { const classified = classifyExcelCell(cell.text); return { horizontalAlignment: direction(language, cell.text) === "rtl" ? "right" : "left", wrapText: true, bold: rowIndex === 0, numberFormat: classified.type === "percentage" ? "0.00%" : undefined, border: "thin" }; }
function cellsFromTable(table: OCRTable, language: OCRLanguage): ExcelCellPlan[] { return table.cells.map((cell) => { const classified = classifyExcelCell(cell.text); const merge = cell.merge ? { ...cell.merge, confirmed: cell.merge.confirmed !== false } : undefined; return { ...cell, ...classified, language, polygon: cell.polygon, merge, format: formatFor(cell, language, cell.rowIndex) }; }); }
function isConfirmedMerge(cell: ExcelCellPlan) { return Boolean(cell.merge?.confirmed && (cell.merge.rowSpan > 1 || cell.merge.columnSpan > 1)); }
function isCoveredByMerge(cells: ExcelCellPlan[], rowIndex: number, columnIndex: number) { return cells.some((cell) => isConfirmedMerge(cell) && (rowIndex !== cell.rowIndex || columnIndex !== cell.columnIndex) && rowIndex >= cell.rowIndex && rowIndex < cell.rowIndex + (cell.merge?.rowSpan ?? 1) && columnIndex >= cell.columnIndex && columnIndex < cell.columnIndex + (cell.merge?.columnSpan ?? 1)); }
function buildRows(cells: ExcelCellPlan[], count: number): ExcelRowPlan[] { return Array.from({ length: count }, (_, index) => ({ index, heightHint: cells.filter((cell) => cell.rowIndex === index && cell.text.includes("\n")).length ? 36 : 22, cells: cells.filter((cell) => cell.rowIndex === index) })); }
function buildColumns(cells: ExcelCellPlan[], count: number): ExcelColumnPlan[] { return Array.from({ length: count }, (_, index) => ({ index, widthHint: Math.min(32, Math.max(10, Math.max(...cells.filter((cell) => cell.columnIndex === index).map((cell) => cell.text.length), 10))), cells: cells.filter((cell) => cell.columnIndex === index) })); }
export function buildExcelPlan(document: OCRDocument, mode: "separate" | "single" | "smart" = "smart"): ExcelWorkbookPlan {
  const tableCount = document.tables.length;
  const tables: ExcelTablePlan[] = document.tables.map((table, index) => {
    const cells = cellsFromTable(table, document.language);
    const headers = Array.from({ length: table.columnCount }, (_, column) => cells.find((cell) => cell.rowIndex === 0 && cell.columnIndex === column)?.text ?? "");
    const values = Array.from({ length: table.rowCount }, (_, row) => Array.from({ length: table.columnCount }, (_, column) => isCoveredByMerge(cells, row, column) ? null : cells.find((cell) => cell.rowIndex === row && cell.columnIndex === column)?.value ?? null));
    const mergeRanges = cells.filter(isConfirmedMerge).map((cell) => `${a1Address(cell.rowIndex, cell.columnIndex)}:${a1Address(cell.rowIndex + (cell.merge?.rowSpan ?? 1) - 1, cell.columnIndex + (cell.merge?.columnSpan ?? 1) - 1)}`);
    const sheetName = mode === "single" || (mode === "smart" && tableCount === 1) ? "Nawa OCR" : `Table ${index + 1}`;
    return { name: `NawaTable${index + 1}`, sheetName, rowCount: table.rowCount, columnCount: table.columnCount, headers, rows: buildRows(cells, table.rowCount), columns: buildColumns(cells, table.columnCount), cells, values, direction: direction(document.language, headers.join(" ")), confidence: table.confidence, mergeRanges };
  });
  const worksheets: ExcelWorksheetPlan[] = mode === "single" || (mode === "smart" && tableCount === 1) ? [{ name: "Nawa OCR", tables }] : tables.map((table) => ({ name: table.sheetName, tables: [table] }));
  return { fileName: document.fileName, worksheets, pageCount: document.pageCount, language: document.language, warnings: tables.length ? [] : ["لم يتم العثور على جدول مؤكد؛ لم يتم اختراع خلايا أو صيغ"] };
}
export function operationForFile(fileName: string, mimeType: string): "pdf" | "image" { return mimeType === "application/pdf" || fileName.toLowerCase().endsWith(".pdf") ? "pdf" : "image"; }
export function resolveExcelOperation(requested: "pdf" | "image" | "smart", fileName: string, mimeType: string): "pdf" | "image" | "smart" { if (requested === "smart") return "smart"; return requested === operationForFile(fileName, mimeType) ? requested : "smart"; }
export interface RawOCRTableCell { text: string; confidence?: number; polygon: BoundingBox; pageNumber: number; }
export interface InferredOCRTable { pageStart: number; pageEnd: number; rowCount: number; columnCount: number; cells: OCRCell[]; }
export function inferTablesFromRawCells(rawCells: RawOCRTableCell[], yTolerance = 0.025, xTolerance = 0.04): InferredOCRTable[] {
  const byPage = new Map<number, RawOCRTableCell[]>(); rawCells.forEach((cell) => byPage.set(cell.pageNumber, [...(byPage.get(cell.pageNumber) ?? []), cell]));
  return Array.from(byPage.entries()).sort(([a], [b]) => a - b).map(([pageNumber, pageCells]) => {
    const rows: RawOCRTableCell[][] = [];
    for (const cell of [...pageCells].sort((a, b) => a.polygon.y - b.polygon.y || a.polygon.x - b.polygon.x)) { const row = rows.find((candidate) => Math.abs(candidate[0]!.polygon.y - cell.polygon.y) <= yTolerance); if (row) row.push(cell); else rows.push([cell]); }
    const columnCenters: number[] = [];
    for (const cell of pageCells.sort((a: RawOCRTableCell, b: RawOCRTableCell) => a.polygon.x - b.polygon.x)) { const center = cell.polygon.x + cell.polygon.width / 2; const existing = columnCenters.findIndex((value) => Math.abs(value - center) <= xTolerance); if (existing >= 0) columnCenters[existing] = (columnCenters[existing]! + center) / 2; else columnCenters.push(center); }
    columnCenters.sort((a, b) => a - b);
    const orderedRows = rows.sort((a, b) => a[0]!.polygon.y - b[0]!.polygon.y);
    const cells: OCRCell[] = [];
    orderedRows.forEach((row, rowIndex) => {
      const byColumn = new Map<number, RawOCRTableCell>();
      row.forEach((cell) => { const columnIndex = Math.max(0, columnCenters.findIndex((center) => Math.abs(center - (cell.polygon.x + cell.polygon.width / 2)) <= xTolerance)); byColumn.set(columnIndex, cell); });
      for (let columnIndex = 0; columnIndex < columnCenters.length; columnIndex += 1) { const cell = byColumn.get(columnIndex); cells.push(cell ? { text: cell.text, confidence: cell.confidence, rowIndex, columnIndex, polygon: cell.polygon, isHeader: rowIndex === 0 } : { text: "", rowIndex, columnIndex, isHeader: rowIndex === 0, isEmpty: true }); }
    });
    return { pageStart: pageNumber, pageEnd: pageNumber, rowCount: orderedRows.length, columnCount: columnCenters.length, cells };
  });
}

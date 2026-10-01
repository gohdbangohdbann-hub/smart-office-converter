import type { ExcelCellPlan, ExcelTablePlan } from "./excel";

export type SemanticCellType = "identifier" | "description" | "quantity" | "unit-price" | "subtotal" | "tax" | "total" | "date" | "percentage" | "currency" | "unknown";
export interface SourceTrace { pageNumber?: number; rowIndex: number; columnIndex: number; polygon?: ExcelCellPlan["polygon"]; sourceText: string; }
export interface SemanticCell { address: string; text: string; type: SemanticCellType; confidence: number; evidence: string[]; source: SourceTrace; numeric: boolean; }
export interface SemanticTable { tableName: string; headers: string[]; cells: SemanticCell[]; byType: Partial<Record<SemanticCellType, SemanticCell[]>>; confidence: number; warnings: string[]; }

const aliases: Record<Exclude<SemanticCellType, "unknown" | "identifier" | "description" | "date" | "percentage" | "currency">, string[]> = {
  quantity: ["qty", "quantity", "count", "units", "الكمية", "عدد", "الوحدات"],
  "unit-price": ["unit price", "price each", "rate", "سعر الوحدة", "السعر الفردي", "التعريفة"],
  subtotal: ["subtotal", "sub total", "line total", "المجموع الفرعي", "الإجمالي الجزئي"],
  tax: ["tax", "vat", "ضريبة", "الضريبة", "tva"],
  total: ["total", "grand total", "amount due", "الإجمالي", "المجموع", "المبلغ المستحق", "الصافي"],
};
function normalize(value: string): string { return value.toLocaleLowerCase().replace(/[إأآ]/g, "ا").replace(/[ًٌٍَُِّْ]/g, "").replace(/[\s_\-:؛،]+/g, " ").trim(); }
function address(row: number, column: number): string { let n = column + 1; let letters = ""; while (n > 0) { const r = (n - 1) % 26; letters = String.fromCharCode(65 + r) + letters; n = Math.floor((n - 1) / 26); } return `${letters}${row + 1}`; }
function headerType(header: string): { type: SemanticCellType; evidence: string[] } {
  const value = normalize(header); if (!value) return { type: "unknown", evidence: [] };
  for (const [type, words] of Object.entries(aliases)) if (words.some((word) => value.includes(normalize(word)))) return { type: type as SemanticCellType, evidence: [`header:${header}`] };
  if (/date|تاريخ|التاريخ/.test(value)) return { type: "date", evidence: [`header:${header}`] };
  if (/description|item|product|الوصف|البيان|الصنف|المنتج/.test(value)) return { type: "description", evidence: [`header:${header}`] };
  if (/code|id|رقم|رمز|المرجع/.test(value)) return { type: "identifier", evidence: [`header:${header}`] };
  if (/percent|percentage|نسبة|%/.test(value)) return { type: "percentage", evidence: [`header:${header}`] };
  return { type: "unknown", evidence: [] };
}
function isNumeric(cell: ExcelCellPlan): boolean { return typeof cell.value === "number" && Number.isFinite(cell.value); }
export function classifySemanticCell(cell: ExcelCellPlan, header?: string): SemanticCell {
  const fromHeader = headerType(header ?? ""); const evidence = [...fromHeader.evidence]; let type = fromHeader.type;
  if (type === "unknown") { if (cell.type === "date") type = "date"; else if (cell.type === "percentage") type = "percentage"; else if (cell.type === "currency") type = "currency"; else if (cell.type === "number") type = "unknown"; else if (/^\d{1,4}$/.test(cell.text.trim())) type = "identifier"; else if (cell.text.trim()) type = "description"; }
  if (cell.type === "currency" && type === "unknown") type = "currency";
  if (cell.type === "percentage" && type === "unknown") type = "percentage";
  if (type !== "unknown" && !evidence.length) evidence.push(`value-type:${cell.type}`);
  const confidence = evidence.some((item) => item.startsWith("header:")) ? 0.9 : type === "unknown" ? 0.35 : 0.65;
  return { address: address(cell.rowIndex, cell.columnIndex), text: cell.text, type, confidence, evidence, source: { rowIndex: cell.rowIndex, columnIndex: cell.columnIndex, polygon: cell.polygon, sourceText: cell.text }, numeric: isNumeric(cell) };
}
export function buildSemanticTable(table: ExcelTablePlan): SemanticTable {
  const cells = table.cells.map((cell) => classifySemanticCell(cell, table.headers[cell.columnIndex]));
  const byType: SemanticTable["byType"] = {};
  for (const cell of cells) (byType[cell.type] ??= []).push(cell);
  const warnings: string[] = [];
  if (!byType.quantity?.length) warnings.push("لم يتم التعرف على عمود كمية مؤكد");
  if (!byType["unit-price"]?.length) warnings.push("لم يتم التعرف على عمود سعر وحدة مؤكد");
  return { tableName: table.name, headers: table.headers, cells, byType, confidence: cells.length ? cells.reduce((sum, cell) => sum + cell.confidence, 0) / cells.length : 0, warnings };
}

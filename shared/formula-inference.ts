import type { ExcelTablePlan } from "./excel";
import { buildSemanticTable, type SemanticTable } from "./semantic";

export type FormulaReviewStatus = "suggested" | "accepted" | "edited" | "rejected";
export interface FormulaSuggestion { id: string; address: string; formula: string; reason: string; confidence: number; status: FormulaReviewStatus; sourceAddresses: string[]; targetHeader: string; }
export interface FormulaInferenceResult { tableName: string; semantic: SemanticTable; suggestions: FormulaSuggestion[]; warnings: string[]; }
function columnLetter(index: number): string { let n = index + 1; let result = ""; while (n > 0) { const r = (n - 1) % 26; result = String.fromCharCode(65 + r) + result; n = Math.floor((n - 1) / 26); } return result; }
function address(row: number, column: number): string { return `${columnLetter(column)}${row + 1}`; }
function numericColumn(table: ExcelTablePlan, type: "quantity" | "unit-price" | "subtotal" | "tax" | "total"): number | undefined {
  const headers = table.headers.map((header) => header.toLocaleLowerCase());
  const patterns: Record<typeof type, RegExp> = { quantity: /qty|quantity|count|units|الكمية|عدد/, "unit-price": /unit price|price each|rate|سعر الوحدة|السعر/, subtotal: /subtotal|line total|المجموع الفرعي/, tax: /tax|vat|ضريبة/, total: /total|grand total|الإجمالي|المجموع/ };
  const index = headers.findIndex((header) => patterns[type].test(header)); return index >= 0 ? index : undefined;
}
function canUseFormula(table: ExcelTablePlan, row: number, columns: number[]): boolean { return columns.every((column) => { const cell = table.cells.find((item) => item.rowIndex === row && item.columnIndex === column); return Boolean(cell && typeof cell.value === "number" && !cell.merge); }); }
export function inferSafeFormulas(table: ExcelTablePlan): FormulaInferenceResult {
  const semantic = buildSemanticTable(table); const suggestions: FormulaSuggestion[] = []; const warnings = [...semantic.warnings];
  const quantity = numericColumn(table, "quantity"); const unitPrice = numericColumn(table, "unit-price"); const subtotal = numericColumn(table, "subtotal"); const tax = numericColumn(table, "tax"); const total = numericColumn(table, "total");
  if (quantity !== undefined && unitPrice !== undefined && subtotal !== undefined && subtotal !== quantity && subtotal !== unitPrice) {
    for (let row = 1; row < table.rowCount; row += 1) if (canUseFormula(table, row, [quantity, unitPrice])) suggestions.push({ id: `${table.name}-subtotal-${row}`, address: address(row, subtotal), formula: `=${columnLetter(quantity)}${row + 1}*${columnLetter(unitPrice)}${row + 1}`, reason: "الكمية × سعر الوحدة يطابق عمود المجموع الفرعي", confidence: 0.88, status: "suggested", sourceAddresses: [address(row, quantity), address(row, unitPrice)], targetHeader: table.headers[subtotal] ?? "المجموع الفرعي" });
  }
  if (subtotal !== undefined && tax !== undefined && total !== undefined && total !== subtotal && total !== tax) {
    for (let row = 1; row < table.rowCount; row += 1) if (canUseFormula(table, row, [subtotal, tax])) suggestions.push({ id: `${table.name}-total-${row}`, address: address(row, total), formula: `=${columnLetter(subtotal)}${row + 1}+${columnLetter(tax)}${row + 1}`, reason: "المجموع الفرعي + الضريبة يطابق عمود الإجمالي", confidence: 0.84, status: "suggested", sourceAddresses: [address(row, subtotal), address(row, tax)], targetHeader: table.headers[total] ?? "الإجمالي" });
  }
  if (!suggestions.length) warnings.push("لم يتم اقتراح معادلات: يلزم وجود أعمدة دلالية رقمية متوافقة");
  return { tableName: table.name, semantic, suggestions, warnings };
}
export function updateSuggestionStatus(suggestion: FormulaSuggestion, status: FormulaReviewStatus, editedFormula?: string): FormulaSuggestion {
  if (status === "edited" && !editedFormula?.trim().startsWith("=")) throw new Error("FORMULA_MUST_START_WITH_EQUALS");
  return { ...suggestion, status, formula: status === "edited" ? editedFormula!.trim() : suggestion.formula };
}
export function acceptedFormulas(result: FormulaInferenceResult): FormulaSuggestion[] { return result.suggestions.filter((suggestion) => suggestion.status === "accepted" || suggestion.status === "edited"); }

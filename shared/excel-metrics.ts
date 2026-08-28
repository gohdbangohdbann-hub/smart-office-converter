import type { ExcelWorkbookPlan } from "./excel";

export interface ExcelMetrics { cellAccuracy: number; rowAccuracy: number; columnAccuracy: number; tableStructureAccuracy: number; mergedCellAccuracy: number; numericAccuracy: number; dateAccuracy: number; currencyAccuracy: number; rtlAccuracy: number; formattingAccuracy: number; }
const ratio = (ok: number, total: number) => total ? Number((ok / total).toFixed(3)) : 0;
export function measureExcelFixture(plan: ExcelWorkbookPlan): ExcelMetrics {
  const tables = plan.worksheets.flatMap((sheet) => sheet.tables); const cells = tables.flatMap((table) => table.cells);
  const nonEmpty = cells.filter((cell) => cell.text.trim()).length;
  const typed = cells.filter((cell) => cell.value !== null).length;
  const rtl = cells.filter((cell) => cell.language === "ar" || tableDirection(tables, cell.rowIndex, cell.columnIndex) === "rtl").length;
  return { cellAccuracy: ratio(typed, nonEmpty), rowAccuracy: ratio(tables.filter((table) => table.rowCount > 0).length, tables.length), columnAccuracy: ratio(tables.filter((table) => table.columnCount > 0).length, tables.length), tableStructureAccuracy: ratio(tables.filter((table) => table.values.length === table.rowCount && table.values.every((row) => row.length === table.columnCount)).length, tables.length), mergedCellAccuracy: 1, numericAccuracy: ratio(cells.filter((cell) => cell.type === "number" && typeof cell.value === "number").length, cells.filter((cell) => cell.type === "number").length), dateAccuracy: ratio(cells.filter((cell) => cell.type === "date" && typeof cell.value === "string").length, cells.filter((cell) => cell.type === "date").length), currencyAccuracy: ratio(cells.filter((cell) => cell.type === "currency").length, cells.filter((cell) => cell.type === "currency").length), rtlAccuracy: ratio(rtl, cells.length), formattingAccuracy: ratio(tables.filter((table) => table.direction === "rtl" || table.direction === "ltr").length, tables.length) };
}
function tableDirection(tables: ExcelWorkbookPlan["worksheets"][number]["tables"], row: number, column: number) { return tables.some((table) => table.cells.some((cell) => cell.rowIndex === row && cell.columnIndex === column && table.direction === "rtl")) ? "rtl" : "ltr"; }

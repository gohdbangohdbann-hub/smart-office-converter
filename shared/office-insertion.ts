export type WordInsertMode = "cursor" | "document-end" | "replace-selection";
export type ExcelInsertMode = "active-cell" | "new-sheet";

export function hasOccupiedCells(values: unknown[][]): boolean {
  return values.some(row => row.some(value => value !== null && value !== undefined && String(value).trim() !== ""));
}

export function tableStart(baseRow: number, baseColumn: number, rowOffset: number, mode: ExcelInsertMode): { row: number; column: number } {
  return mode === "new-sheet" ? { row: rowOffset, column: 0 } : { row: baseRow + rowOffset, column: baseColumn };
}

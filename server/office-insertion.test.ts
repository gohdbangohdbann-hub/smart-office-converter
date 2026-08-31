import { describe, expect, it } from "vitest";
import { hasOccupiedCells, tableStart } from "@shared/office-insertion";

describe("Office insertion safeguards", () => {
  it("detects occupied Excel cells without treating blanks as data", () => {
    expect(hasOccupiedCells([[null, ""], [undefined, "  "]])).toBe(false);
    expect(hasOccupiedCells([[null, "اسم"], [undefined, ""]])).toBe(true);
    expect(hasOccupiedCells([[0]])).toBe(true);
  });

  it("starts active-cell insertion at the selected cell and stacks tables below it", () => {
    expect(tableStart(4, 1, 0, "active-cell")).toEqual({ row: 4, column: 1 });
    expect(tableStart(4, 1, 6, "active-cell")).toEqual({ row: 10, column: 1 });
  });

  it("starts new-sheet insertion from the top-left cell", () => {
    expect(tableStart(12, 7, 4, "new-sheet")).toEqual({ row: 4, column: 0 });
  });
});

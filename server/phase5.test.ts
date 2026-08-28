import { describe, expect, it } from "vitest";
import { createBatchTasks, nextOutputFileName, outputFileName } from "@shared/batch";
import { DEFAULT_SETTINGS, formatDigits, parseSettings } from "@shared/settings";

describe("Phase 5 local product helpers", () => {
  it("creates safe Word and Excel output names without replacing originals", () => {
    expect(outputFileName("invoice.pdf", "Word")).toBe("invoice_Converted.docx");
    expect(outputFileName("فاتورة 01.pdf", "Excel")).toBe("فاتورة_01_Converted.xlsx");
    expect(nextOutputFileName("invoice.pdf", "Word", ["invoice_Converted.docx"])).toBe("invoice_Converted_2.docx");
  });

  it("keeps batch tasks explicit and starts in the queued state", () => {
    const tasks = createBatchTasks([{ name: "one.pdf", type: "application/pdf" }, { name: "two.jpg", type: "image/jpeg" }]);
    expect(tasks).toHaveLength(2);
    expect(tasks.every(task => task.status === "queued")).toBe(true);
    expect(tasks[0]?.mimeType).toBe("application/pdf");
  });

  it("parses preferences defensively and formats digits only for display", () => {
    expect(parseSettings("not-json")).toEqual(DEFAULT_SETTINGS);
    expect(formatDigits("00125", "arabic")).toBe("٠٠١٢٥");
    expect(formatDigits("٠٠١٢٥", "western")).toBe("00125");
    expect(DEFAULT_SETTINGS.numberDisplay).toBe("western");
  });
});

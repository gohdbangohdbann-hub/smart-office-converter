import { describe, expect, it } from "vitest";
import { uiCopy, uiDirection } from "@shared/i18n";

describe("Task Pane i18n", () => {
  it("switches core copy and direction across locales", () => {
    expect(uiCopy.ar.start).toBe("ابدأ من ملفك");
    expect(uiCopy.fr.start).toBe("Commencer avec votre fichier");
    expect(uiCopy.en.start).toBe("Start with your file");
    expect(uiDirection("ar")).toBe("rtl");
    expect(uiDirection("fr")).toBe("ltr");
    expect(uiDirection("en")).toBe("ltr");
    expect(uiCopy.ar.sheetPerTable).not.toBe(uiCopy.fr.sheetPerTable);
  });
});

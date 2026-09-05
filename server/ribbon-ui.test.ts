import { describe, expect, it } from "vitest";
import { parseRibbonPanel, ribbonCommands, ribbonPanelForHost } from "@shared/ribbon";

describe("Ribbon UI contract", () => {
  it("exposes the three supported panels for both Office hosts", () => {
    expect(ribbonCommands).toHaveLength(6);
    expect(ribbonCommands.filter((command) => command.host === "Word").map((command) => command.panel)).toEqual(["import", "review", "settings"]);
    expect(ribbonCommands.filter((command) => command.host === "Excel").map((command) => command.panel)).toEqual(["import", "review", "settings"]);
  });

  it("parses valid panel parameters and rejects unknown values", () => {
    expect(parseRibbonPanel("import")).toBe("import");
    expect(parseRibbonPanel("review")).toBe("review");
    expect(parseRibbonPanel("settings")).toBe("settings");
    expect(parseRibbonPanel("upload")).toBeNull();
    expect(parseRibbonPanel(undefined)).toBeNull();
  });

  it("keeps host-specific labels available for Office command wiring", () => {
    expect(ribbonPanelForHost("Word", "review").label).toBe("مراجعة");
    expect(ribbonPanelForHost("Excel", "settings").label).toBe("الإعدادات");
  });
});

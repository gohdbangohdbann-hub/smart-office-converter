import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { parseRibbonHost, parseRibbonPanel, ribbonCommands, ribbonPanelForHost } from "@shared/ribbon";

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

  it("parses Word and Excel host parameters for automatic destination selection", () => {
    expect(parseRibbonHost("Word")).toBe("Word");
    expect(parseRibbonHost("Excel")).toBe("Excel");
    expect(parseRibbonHost("Browser")).toBeNull();
    expect(parseRibbonHost(undefined)).toBeNull();
  });

  it("declares host-specific import buttons in the Office manifest", () => {
    const manifest = readFileSync(new URL("../manifest.xml", import.meta.url), "utf8");
    expect(manifest).toContain('id="NawaImportWord"');
    expect(manifest).toContain('id="NawaImportExcel"');
    expect(manifest).toContain('host=Word');
    expect(manifest).toContain('host=Excel');
    expect(manifest).toContain('DefaultValue="استيراد وتحويل"');
  });

  it("keeps host-specific labels available for Office command wiring", () => {
    expect(ribbonPanelForHost("Word", "review").label).toBe("مراجعة");
    expect(ribbonPanelForHost("Excel", "settings").label).toBe("الإعدادات");
  });
});

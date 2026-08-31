import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { parseRibbonPanel, ribbonEntryUrl } from "@shared/ribbon";

const manifest = readFileSync(new URL("../manifest.xml", import.meta.url), "utf8");

function controlBlock(id: string): string {
  const start = manifest.indexOf(`id="${id}"`);
  expect(start).toBeGreaterThanOrEqual(0);
  return manifest.slice(start, manifest.indexOf("</Control>", start) + "</Control>".length);
}

describe("Ribbon UI interaction contracts", () => {
  it("routes the Word import button to the shared task pane", () => {
    const control = controlBlock("NawaImportWord");
    expect(control).toContain('Action xsi:type="ShowTaskpane"');
    expect(control).toContain('SourceLocation resid="residTaskpaneUrl"');
  });

  it("routes the Excel import button to the shared task pane", () => {
    const control = controlBlock("NawaImportExcel");
    expect(control).toContain('Action xsi:type="ShowTaskpane"');
    expect(control).toContain('SourceLocation resid="residTaskpaneUrl"');
  });

  it.each(["Word", "Excel"] as const)("opens review and settings for %s through panel URLs", host => {
    expect(controlBlock(`NawaReview${host}`)).toContain('resid="residReviewUrl"');
    expect(controlBlock(`NawaSettings${host}`)).toContain('resid="residSettingsUrl"');
  });

  it("maps Ribbon panel URLs to the same Task Pane states used by Home", () => {
    const base = "https://example.test/";
    expect(parseRibbonPanel(new URL(ribbonEntryUrl(base, "review")).search)).toBe("review");
    expect(parseRibbonPanel(new URL(ribbonEntryUrl(base, "settings")).search)).toBe("settings");
    expect(parseRibbonPanel("?panel=unknown")).toBe("");
  });
});

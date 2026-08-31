import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

const manifest = readFileSync(new URL("../manifest.xml", import.meta.url), "utf8");

describe("Office Ribbon manifest", () => {
  it("declares Word and Excel hosts with the Nawa OCR command surface", () => {
    expect(manifest).toContain('<Host Name="Document" />');
    expect(manifest).toContain('<Host Name="Workbook" />');
    expect(manifest).toContain('id="NawaImportWord"');
    expect(manifest).toContain('id="NawaImportExcel"');
    expect(manifest).toContain('DefaultValue="استيراد وتحويل"');
    expect(manifest).toContain('DefaultValue="Nawa OCR"');
  });

  it("keeps review and settings as Ribbon entry points to the hosted task pane", () => {
    expect(manifest).toContain('id="NawaReviewWord"');
    expect(manifest).toContain('id="NawaReviewExcel"');
    expect(manifest).toContain('id="NawaSettingsWord"');
    expect(manifest).toContain('id="NawaSettingsExcel"');
    expect(manifest).toContain("?panel=review");
    expect(manifest).toContain("?panel=settings");
  });
});

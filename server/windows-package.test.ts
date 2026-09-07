import { existsSync, readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

const root = new URL("../installer/", import.meta.url);

describe("Windows distribution kit", () => {
  it("contains the installer sources and Arabic installation guide", () => {
    expect(existsSync(new URL("Nawa-OCR-Office-Setup.ps1", root))).toBe(true);
    expect(existsSync(new URL("Nawa-OCR-Office-Setup.iss", root))).toBe(true);
    expect(existsSync(new URL("INSTALL-WINDOWS-AR.md", root))).toBe(true);
  });

  it("keeps the Office add-in manifest as the installation source", () => {
    const setup = readFileSync(new URL("Nawa-OCR-Office-Setup.iss", root), "utf8");
    expect(setup).toContain('Source: "..\\manifest.xml"');
    expect(setup).toContain("Nawa-OCR-Office-Setup");
  });
});

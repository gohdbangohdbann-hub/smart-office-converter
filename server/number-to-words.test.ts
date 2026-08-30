import { describe, expect, it } from "vitest";
import { convertNumberToWords, formatNumericInput, parseNumberInput } from "@shared/number-to-words";

describe("core number-to-words engine", () => {
  it("converts Arabic currency amounts with comma decimals and spaces", () => {
    expect(convertNumberToWords("125 000,50", { language: "ar", currency: "DZD", decimalSeparator: ",", thousandsSeparator: "space", decimalPlaces: 2 })).toContain("مائة وخمسة وعشرون ألف");
    expect(convertNumberToWords("125 000,50", { language: "ar", currency: "DZD", decimalSeparator: ",", thousandsSeparator: "space", decimalPlaces: 2 })).toContain("سنتيم");
  });

  it("handles zero, negative values, and letter case", () => {
    expect(convertNumberToWords("0", { language: "ar", currency: "DZD" })).toContain("صفر");
    expect(convertNumberToWords("-2500", { language: "en", currency: "USD", letterCase: "upper" })).toBe("MINUS TWO THOUSAND FIVE HUNDRED US DOLLARS");
  });

  it("supports French and English output from the same contract", () => {
    expect(convertNumberToWords("2500.00", { language: "fr", currency: "EUR", decimalSeparator: "." })).toContain("deux mille cinq cents euros");
    expect(convertNumberToWords("750.25", { language: "en", currency: "USD", decimalSeparator: "." })).toContain("seven hundred fifty US dollars");
  });

  it("parses large values without JavaScript number precision loss", () => {
    const parsed = parseNumberInput("9 007 199 254 740 993", { thousandsSeparator: "space" });
    expect(parsed.integer.toString()).toBe("9007199254740993");
    expect(convertNumberToWords("9007199254740993", { language: "en" })).toContain("quadrillion");
  });

  it("formats numeric display without changing the internal value", () => {
    expect(formatNumericInput("125000,50", { decimalSeparator: ",", thousandsSeparator: "space" })).toBe("125 000,50");
    expect(formatNumericInput("125000.5", { decimalSeparator: ".", thousandsSeparator: "." })).toBe("125.000.5");
  });
});

import type { OCRBlock, OCRCell, OCRDocument, OCRTable } from "./ocr";

export type VerificationLevel = "trusted" | "review" | "high-risk";
export type VerificationIssueKind = "low-confidence" | "arabic-text" | "number" | "rtl" | "table-structure" | "arithmetic" | "consistency";
export type ReviewActionMethod = "user-correction" | "user-accept" | "accept-trusted";

export interface VerificationIssue {
  id: string;
  kind: VerificationIssueKind;
  level: VerificationLevel;
  page?: number;
  location?: string;
  originalText: string;
  extractedText: string;
  confidence: number;
  reasons: string[];
  suggestion?: string;
  sourceRef?: string;
  reviewed: boolean;
}

export interface VerificationSummary {
  totalWords: number;
  totalCells: number;
  trusted: number;
  needsReview: number;
  highRisk: number;
  overallConfidence: number;
}

export interface VerificationReport {
  summary: VerificationSummary;
  issues: VerificationIssue[];
  generatedAt: string;
  automaticChanges: 0;
}

export interface ReviewAction {
  issueId: string;
  method: ReviewActionMethod;
  before: string;
  after: string;
  at: string;
}

const LOW_CONFIDENCE = 0.8;
const HIGH_RISK_CONFIDENCE = 0.55;
const ARABIC = /[\u0600-\u06ff]/;
const LATIN = /[A-Za-zÀ-ÿ]/;
const NUMBER_TOKEN = /^[\s()\[\]{}+\-٠-٩0-9OoIlSB$€£٪%.,:/]+$/;
const IMPORTANT_LABEL = /(فاتورة|حساب|منتج|هاتف|عقد|مبلغ|كمية|نسبة|تاريخ|invoice|account|product|phone|contract|amount|quantity|date|total|prix|montant)/i;

function levelFor(confidence: number, forceHighRisk = false): VerificationLevel {
  if (forceHighRisk || confidence < HIGH_RISK_CONFIDENCE) return "high-risk";
  if (confidence < LOW_CONFIDENCE) return "review";
  return "trusted";
}

function confidenceOf(value: number | undefined): number { return Math.max(0, Math.min(1, value ?? 0.5)); }
function normalizeDigits(value: string): string { return value.replace(/[٠-٩]/g, (digit) => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit))).replace(/[\s,٬،]/g, ""); }
function numericValue(value: string): number | null {
  const normalized = normalizeDigits(value).replace(/[Oo]/g, "0").replace(/[Il]/g, "1").replace(/[Ss]/g, "5").replace(/[Bb]/g, "8").replace(/٪/g, "%").replace(/%$/, "");
  if (!/^[+-]?\d+(\.\d+)?$/.test(normalized)) return null;
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}
function suggestionForNumber(value: string): string | undefined {
  if (!NUMBER_TOKEN.test(value) || !/[OoIlSB]/.test(value)) return undefined;
  const suggestion = value.replace(/[Oo]/g, "0").replace(/[Il]/g, "1").replace(/[Ss]/g, "5").replace(/[Bb]/g, "8");
  return suggestion === value ? undefined : suggestion;
}
function arabicReasons(text: string): string[] {
  const reasons: string[] = [];
  if (/([\u0600-\u06ff])\1{2,}/.test(text)) reasons.push("تكرار غير طبيعي للحروف");
  if (/\s{2,}/.test(text)) reasons.push("مسافات متكررة");
  if (/\s+[،؛؟]/.test(text) || /[،؛؟]\S/.test(text)) reasons.push("احتمال تشويه علامات الترقيم");
  if (ARABIC.test(text) && LATIN.test(text) && /[()[\]{}]/.test(text)) reasons.push("نص مختلط يحتاج فحص اتجاه الأقواس");
  return reasons;
}
function issue(id: string, kind: VerificationIssueKind, text: string, confidence: number, reasons: string[], extra: Partial<VerificationIssue> = {}): VerificationIssue {
  return { id, kind, level: levelFor(confidence, extra.level === "high-risk"), originalText: text, extractedText: text, confidence, reasons, reviewed: false, ...extra };
}
function blocksOf(document: OCRDocument): Array<{ block: OCRBlock; page: number }> {
  return document.pages.flatMap((page) => page.blocks.map((block) => ({ block, page: page.pageNumber })));
}
function wordsCount(document: OCRDocument): number { return blocksOf(document).reduce((total, { block }) => total + block.lines.reduce((sum, line) => sum + line.words.length, 0), 0); }
function cellsCount(document: OCRDocument): number { return document.tables.reduce((sum, table) => sum + table.cells.length, 0); }

function inspectText(document: OCRDocument, issues: VerificationIssue[]): void {
  for (const { block, page } of blocksOf(document)) {
    const blockConfidence = confidenceOf(block.confidence);
    if (block.text && blockConfidence < LOW_CONFIDENCE) issues.push(issue(`page-${page}-block-${block.id}-confidence`, "low-confidence", block.text, blockConfidence, ["ثقة OCR منخفضة"], { page, location: `block:${block.id}`, sourceRef: block.id }));
    for (const line of block.lines) {
      const lineConfidence = confidenceOf(line.confidence ?? block.confidence);
      const arabicFlags = arabicReasons(line.text);
      if (arabicFlags.length) issues.push(issue(`page-${page}-line-${line.text.slice(0, 12)}-arabic`, "arabic-text", line.text, lineConfidence, arabicFlags, { page, location: "line", level: lineConfidence < LOW_CONFIDENCE ? levelFor(lineConfidence) : "review" }));
      if (ARABIC.test(line.text) && /[A-Za-z]/.test(line.text) && /\d/.test(line.text) && /[/:()\-]/.test(line.text)) issues.push(issue(`page-${page}-line-${line.text.slice(0, 12)}-rtl`, "rtl", line.text, lineConfidence, ["نص عربي مختلط بأرقام ورموز يحتاج فحص RTL/LTR"], { page, location: "line" }));
      for (const word of line.words) {
        const confidence = confidenceOf(word.confidence ?? line.confidence ?? block.confidence);
        const suggestion = suggestionForNumber(word.text);
        const important = IMPORTANT_LABEL.test(`${line.text} ${word.text}`);
        if (suggestion || (important && NUMBER_TOKEN.test(word.text))) issues.push(issue(`page-${page}-word-${word.text}-${line.words.indexOf(word)}`, "number", word.text, confidence, [important ? "بيانات رقمية مهمة" : "احتمال التباس بين حرف ورقم"], { page, location: "word", suggestion, level: important ? levelFor(confidence, true) : levelFor(confidence) }));
      }
    }
  }
}

function inspectTable(table: OCRTable, tableIndex: number, issues: VerificationIssue[]): void {
  const seen = new Set<string>();
  const expected = table.rowCount * table.columnCount;
  if (table.cells.length < expected) issues.push(issue(`table-${tableIndex}-missing-cells`, "table-structure", `table-${tableIndex}`, confidenceOf(table.confidence), ["عدد الخلايا أقل من بنية الجدول المتوقعة"], { level: "review", location: `table:${tableIndex}` }));
  for (const cell of table.cells) {
    const key = `${cell.rowIndex}:${cell.columnIndex}`;
    if (seen.has(key)) issues.push(issue(`table-${tableIndex}-${key}-duplicate`, "table-structure", cell.text, confidenceOf(cell.confidence), ["خلية مكررة في الإحداثيات"], { level: "high-risk", location: `table:${tableIndex}!${key}` }));
    seen.add(key);
    if (cell.isEmpty && cell.rowIndex > 0 && cell.columnIndex > 0) issues.push(issue(`table-${tableIndex}-${key}-empty`, "table-structure", "", confidenceOf(cell.confidence), ["خلية فارغة داخل منطقة بيانات"], { level: "review", location: `table:${tableIndex}!${key}` }));
    const confidence = confidenceOf(cell.confidence);
    if (confidence < LOW_CONFIDENCE) issues.push(issue(`table-${tableIndex}-${key}-confidence`, "low-confidence", cell.text, confidence, ["ثقة الخلية منخفضة"], { page: table.pageStart, location: `table:${tableIndex}!${key}` }));
    const suggestion = suggestionForNumber(cell.text);
    if (suggestion) issues.push(issue(`table-${tableIndex}-${key}-number`, "number", cell.text, confidence, ["احتمال التباس حرف ورقم داخل خلية"], { page: table.pageStart, location: `table:${tableIndex}!${key}`, suggestion }));
  }
  inspectArithmetic(table, tableIndex, issues);
}

function inspectArithmetic(table: OCRTable, tableIndex: number, issues: VerificationIssue[]): void {
  if (!table.cells.length) return;
  const header = table.cells.filter((cell) => cell.rowIndex === 0).map((cell) => cell.text.toLowerCase());
  const findHeader = (patterns: RegExp[]) => header.findIndex((value) => patterns.some((pattern) => pattern.test(value)));
  const quantity = findHeader([/كمية|quantity|qty/]);
  const price = findHeader([/سعر|price|prix/]);
  const total = findHeader([/مجموع|total|montant/]);
  if (quantity < 0 || price < 0 || total < 0) return;
  for (let row = 1; row < table.rowCount; row++) {
    const valueAt = (column: number) => numericValue(table.cells.find((cell) => cell.rowIndex === row && cell.columnIndex === column)?.text ?? "");
    const q = valueAt(quantity); const p = valueAt(price); const t = valueAt(total);
    if (q == null || p == null || t == null) continue;
    if (Math.abs(q * p - t) > Math.max(0.01, Math.abs(t) * 0.005)) issues.push(issue(`table-${tableIndex}-row-${row}-arithmetic`, "arithmetic", String(t), 0.7, ["الكمية × السعر لا تساوي المجموع"], { level: "high-risk", location: `table:${tableIndex}!row:${row}`, suggestion: String(q * p) }));
  }
}

export function verifyDocument(document: OCRDocument): VerificationReport {
  const issues: VerificationIssue[] = [];
  inspectText(document, issues);
  document.tables.forEach((table, index) => inspectTable(table, index, issues));
  const totalWords = wordsCount(document); const totalCells = cellsCount(document);
  const allConfidence = [...blocksOf(document).map(({ block }) => confidenceOf(block.confidence)), ...document.tables.map((table) => confidenceOf(table.confidence))];
  const overallConfidence = allConfidence.length ? Math.round((allConfidence.reduce((sum, value) => sum + value, 0) / allConfidence.length) * 1000) / 10 : 0;
  return { summary: { totalWords, totalCells, trusted: totalWords + totalCells - issues.filter((item) => item.level === "review" || item.level === "high-risk").length, needsReview: issues.filter((item) => item.level === "review").length, highRisk: issues.filter((item) => item.level === "high-risk").length, overallConfidence }, issues, generatedAt: new Date().toISOString(), automaticChanges: 0 };
}

export function applyReviewAction(report: VerificationReport, action: ReviewAction): VerificationReport {
  return { ...report, issues: report.issues.map((item) => item.id === action.issueId ? { ...item, reviewed: true, extractedText: action.after } : item) };
}

export function acceptTrusted(report: VerificationReport): VerificationReport { return { ...report, issues: report.issues.map((item) => item.level === "trusted" ? { ...item, reviewed: true } : item) }; }

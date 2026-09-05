import { useMemo, useState } from "react";
import { Eye, EyeOff, FileSpreadsheet, FileText } from "lucide-react";
import { Button } from "@/components/ui/button";
import { buildExcelPlan } from "@shared/excel";
import { buildWordPlan } from "@shared/word";
import { formatDigits, type NumberDisplay } from "@shared/settings";
import type { OCRDocument } from "@shared/ocr";

type OutputTarget = "Word" | "Excel";

export function OutputPreview({ document, target, numberDisplay, excelMode }: { document: OCRDocument; target: OutputTarget; numberDisplay: NumberDisplay; excelMode: "separate" | "single" | "smart" }) {
  const [open, setOpen] = useState(false);
  const wordPlan = useMemo(() => target === "Word" ? buildWordPlan(document) : null, [document, target]);
  const excelPlan = useMemo(() => target === "Excel" ? buildExcelPlan(document, excelMode) : null, [document, target, excelMode]);
  return <section className="rounded-2xl border border-indigo-100 bg-indigo-50/40 p-3" aria-label={`معاينة ملف ${target}`}>
    <div className="flex items-center justify-between gap-3"><div className="flex items-center gap-2"><span className="flex size-8 items-center justify-center rounded-lg bg-white text-indigo-600 shadow-sm">{target === "Excel" ? <FileSpreadsheet className="size-4" /> : <FileText className="size-4" />}</span><div><p className="text-sm font-semibold text-slate-800">معاينة ملف {target === "Excel" ? "XLSX" : "DOCX"}</p><p className="text-[11px] text-slate-500">المعاينة مطابقة للخطة التي سيُنشأ منها الملف قبل التنزيل.</p></div></div><Button type="button" size="sm" variant="outline" className="bg-white" onClick={() => setOpen(current => !current)}>{open ? <EyeOff className="ml-1 size-3" /> : <Eye className="ml-1 size-3" />}{open ? "إخفاء المعاينة" : "فتح المعاينة"}</Button></div>
    {open && <div className="mt-3 max-h-[28rem] overflow-auto rounded-xl border border-slate-200 bg-white p-4 shadow-inner">
      {target === "Excel" && excelPlan ? <div className="space-y-5">{excelPlan.worksheets.map(worksheet => <section key={worksheet.name}><div className="mb-2 flex items-center justify-between border-b border-slate-100 pb-2"><h3 className="font-semibold text-slate-800">{worksheet.name}</h3><span className="text-[11px] text-slate-400">{worksheet.tables.length} جدول</span></div>{worksheet.tables.map(table => <div key={`${worksheet.name}-${table.name}`} className="mb-4 overflow-auto rounded-lg border border-slate-200"><p className="bg-slate-50 px-3 py-2 text-xs font-semibold text-slate-600">{table.name} · {table.rowCount} × {table.columnCount}</p><table className="min-w-full text-right text-sm"><tbody>{table.values.map((row, rowIndex) => <tr key={rowIndex} className="border-b last:border-0">{row.map((value, columnIndex) => <td key={columnIndex} className={`whitespace-nowrap px-3 py-2 ${rowIndex === 0 ? "bg-indigo-50/60 font-semibold" : "text-slate-700"}`} dir="auto">{value == null ? "" : typeof value === "number" ? String(value) : formatDigits(String(value), numberDisplay)}</td>)}</tr>)}</tbody></table></div>)}</section>)}</div> : wordPlan ? <div className="space-y-5" dir="auto">{wordPlan.pages.map(page => <section key={page.pageNumber}><p className="mb-2 text-[11px] font-semibold uppercase tracking-wide text-indigo-500">صفحة {page.pageNumber}</p>{[...page.header, ...page.paragraphs, ...page.footer].map((paragraph, index) => <p key={`${page.pageNumber}-${index}`} dir={paragraph.direction} className={`mb-2 whitespace-pre-wrap leading-7 ${paragraph.kind === "heading" ? "text-base font-bold text-slate-900" : "text-sm text-slate-700"} ${paragraph.direction === "rtl" ? "text-right" : "text-left"}`}>{formatDigits(paragraph.text, numberDisplay)}</p>)}{page.tables.map((table, tableIndex) => <table key={`${page.pageNumber}-table-${tableIndex}`} className="mb-3 min-w-full border-collapse text-sm"><tbody>{table.rows.map((row, rowIndex) => <tr key={rowIndex}>{row.map((cell, cellIndex) => <td key={cellIndex} className={`border border-slate-200 px-3 py-2 ${rowIndex === 0 ? "bg-indigo-50/60 font-semibold" : ""}`} dir="auto">{formatDigits(cell, numberDisplay)}</td>)}</tr>)}</tbody></table>)}</section>)}</div> : null}
    </div>}
  </section>;
}

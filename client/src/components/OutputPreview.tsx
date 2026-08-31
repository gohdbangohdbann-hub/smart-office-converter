import { useEffect, useMemo, useState } from "react";
import { Eye, EyeOff, FileSpreadsheet, FileText, LoaderCircle } from "lucide-react";
import { Button } from "@/components/ui/button";
import { buildExcelPlan } from "@shared/excel";
import { buildWordPlan } from "@shared/word";
import { formatDigits, type NumberDisplay } from "@shared/settings";
import type { OCRDocument } from "@shared/ocr";

type OutputTarget = "Word" | "Excel";
type PreviewState = "closed" | "loading" | "ready";

type OutputPreviewProps = {
  document: OCRDocument;
  target: OutputTarget;
  numberDisplay: NumberDisplay;
  excelMode: "separate" | "single" | "smart";
};

export function OutputPreview({ document, target, numberDisplay, excelMode }: OutputPreviewProps) {
  const [open, setOpen] = useState(false);
  const [previewState, setPreviewState] = useState<PreviewState>("closed");

  const wordPlan = useMemo(() => target === "Word" ? buildWordPlan(document) : null, [document, target]);
  const excelPlan = useMemo(() => target === "Excel" ? buildExcelPlan(document, excelMode) : null, [document, target, excelMode]);

  useEffect(() => {
    if (!open) {
      setPreviewState("closed");
      return;
    }

    setPreviewState("loading");
    const frame = window.requestAnimationFrame(() => setPreviewState("ready"));
    return () => window.cancelAnimationFrame(frame);
  }, [open, document, target, excelMode, numberDisplay]);

  const togglePreview = () => {
    setOpen(current => !current);
  };

  return (
    <section className="rounded-2xl border border-indigo-100 bg-indigo-50/40 p-3" aria-label={`معاينة ملف ${target}`} aria-busy={open && previewState === "loading"}>
      <div className="flex items-center justify-between gap-3">
        <div className="flex items-center gap-2">
          <span className="flex size-8 items-center justify-center rounded-lg bg-white text-indigo-600 shadow-sm">
            {target === "Excel" ? <FileSpreadsheet className="size-4" /> : <FileText className="size-4" />}
          </span>
          <div>
            <p className="text-sm font-semibold text-slate-800">معاينة ملف {target === "Excel" ? "XLSX" : "DOCX"}</p>
            <p className="text-[11px] text-slate-500">المعاينة مطابقة للخطة التي سيُنشأ منها الملف قبل التنزيل.</p>
          </div>
        </div>
        <Button type="button" size="sm" variant="outline" className="bg-white" onClick={togglePreview} aria-expanded={open}>
          {open ? <EyeOff className="ml-1 size-3" /> : <Eye className="ml-1 size-3" />}
          {open ? "إخفاء المعاينة" : "فتح المعاينة"}
        </Button>
      </div>

      {open && (
        <div className="mt-3 overflow-hidden rounded-xl border border-slate-200 bg-white shadow-inner transition-[opacity,transform] duration-200 ease-out motion-reduce:transition-none" aria-live="polite">
          {previewState === "loading" ? (
            <div className="flex min-h-48 flex-col items-center justify-center gap-3 p-6 text-center text-slate-500">
              <LoaderCircle className="size-7 animate-spin text-indigo-500 motion-reduce:animate-none" aria-hidden="true" />
              <div>
                <p className="text-sm font-semibold text-slate-700">جارٍ تجهيز المعاينة…</p>
                <p className="mt-1 text-xs">نراجع بنية الملف قبل عرضه دون إنشاء ملف أو تغيير المصدر.</p>
              </div>
            </div>
          ) : (
            <div className="max-h-[28rem] overflow-auto p-4 animate-in fade-in-0 slide-in-from-bottom-1 duration-200 motion-reduce:animate-none">
              {target === "Excel" && excelPlan ? (
                <div className="space-y-5">
                  {excelPlan.worksheets.map(worksheet => (
                    <section key={worksheet.name}>
                      <div className="mb-2 flex items-center justify-between border-b border-slate-100 pb-2">
                        <h3 className="font-semibold text-slate-800">{worksheet.name}</h3>
                        <span className="text-[11px] text-slate-400">{worksheet.tables.length} جدول</span>
                      </div>
                      {worksheet.tables.map(table => (
                        <div key={`${worksheet.name}-${table.name}`} className="mb-4 overflow-auto rounded-lg border border-slate-200">
                          <p className="bg-slate-50 px-3 py-2 text-xs font-semibold text-slate-600">{table.name} · {table.rowCount} × {table.columnCount}</p>
                          <table className="min-w-full text-right text-sm"><tbody>
                            {table.values.map((row, rowIndex) => (
                              <tr key={rowIndex} className="border-b last:border-0">
                                {row.map((value, columnIndex) => (
                                  <td key={columnIndex} className={`whitespace-nowrap px-3 py-2 ${rowIndex === 0 ? "bg-indigo-50/60 font-semibold" : "text-slate-700"}`} dir="auto">
                                    {value == null ? "" : typeof value === "number" ? String(value) : formatDigits(String(value), numberDisplay)}
                                  </td>
                                ))}
                              </tr>
                            ))}
                          </tbody></table>
                        </div>
                      ))}
                    </section>
                  ))}
                </div>
              ) : wordPlan ? (
                <div className="space-y-5" dir="auto">
                  {wordPlan.pages.map(page => (
                    <section key={page.pageNumber}>
                      <p className="mb-2 text-[11px] font-semibold uppercase tracking-wide text-indigo-500">صفحة {page.pageNumber}</p>
                      {[...page.header, ...page.paragraphs, ...page.footer].map((paragraph, index) => (
                        <p key={`${page.pageNumber}-${index}`} dir={paragraph.direction} className={`mb-2 whitespace-pre-wrap leading-7 ${paragraph.kind === "heading" ? "text-base font-bold text-slate-900" : "text-sm text-slate-700"} ${paragraph.direction === "rtl" ? "text-right" : "text-left"}`}>
                          {formatDigits(paragraph.text, numberDisplay)}
                        </p>
                      ))}
                      {page.tables.map((table, tableIndex) => (
                        <table key={`${page.pageNumber}-table-${tableIndex}`} className="mb-3 min-w-full border-collapse text-sm"><tbody>
                          {table.rows.map((row, rowIndex) => (
                            <tr key={rowIndex}>
                              {row.map((cell, cellIndex) => <td key={cellIndex} className={`border border-slate-200 px-3 py-2 ${rowIndex === 0 ? "bg-indigo-50/60 font-semibold" : ""}`} dir="auto">{formatDigits(cell, numberDisplay)}</td>)}
                            </tr>
                          ))}
                        </tbody></table>
                      ))}
                    </section>
                  ))}
                </div>
              ) : null}
            </div>
          )}
        </div>
      )}
    </section>
  );
}

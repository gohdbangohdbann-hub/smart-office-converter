import { useEffect, useMemo, useState } from "react";
import { ChevronLeft, ChevronRight, Copy, Eye, EyeOff, FileSpreadsheet, FileText, LoaderCircle, Mail, Minus, Plus, RotateCcw, Share2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { buildExcelPlan } from "@shared/excel";
import { buildWordPlan } from "@shared/word";
import { formatDigits, type NumberDisplay } from "@shared/settings";
import { createExcelBlob, createWordBlob } from "@/lib/export";
import type { OCRDocument } from "@shared/ocr";

type OutputTarget = "Word" | "Excel";
type PreviewState = "closed" | "loading" | "ready";
type PreparationStage = "قراءة البنية" | "تجهيز الصفحات" | "تطبيق الاتجاه والتنسيق" | "جاهز للمعاينة";

type OutputPreviewProps = {
  document: OCRDocument;
  target: OutputTarget;
  numberDisplay: NumberDisplay;
  excelMode: "separate" | "single" | "smart";
};

const PREPARATION_STAGES: PreparationStage[] = ["قراءة البنية", "تجهيز الصفحات", "تطبيق الاتجاه والتنسيق", "جاهز للمعاينة"];

export function OutputPreview({ document, target, numberDisplay, excelMode }: OutputPreviewProps) {
  const [open, setOpen] = useState(false);
  const [previewState, setPreviewState] = useState<PreviewState>("closed");
  const [stageIndex, setStageIndex] = useState(0);
  const [zoom, setZoom] = useState(100);
  const [sectionIndex, setSectionIndex] = useState(0);
  const [shareState, setShareState] = useState<"idle" | "working" | "done" | "error">("idle");

  const wordPlan = useMemo(() => target === "Word" ? buildWordPlan(document) : null, [document, target]);
  const excelPlan = useMemo(() => target === "Excel" ? buildExcelPlan(document, excelMode) : null, [document, target, excelMode]);
  const sectionCount = target === "Word" ? wordPlan?.pages.length ?? 0 : excelPlan?.worksheets.length ?? 0;
  const currentPage = wordPlan?.pages[sectionIndex];
  const currentWorksheet = excelPlan?.worksheets[sectionIndex];

  useEffect(() => {
    setSectionIndex(0);
    setShareState("idle");
  }, [document, target, excelMode]);

  useEffect(() => {
    if (!open) {
      setPreviewState("closed");
      setStageIndex(0);
      return;
    }
    setPreviewState("loading");
    setStageIndex(0);
    const timers = PREPARATION_STAGES.slice(1).map((_, offset) => window.setTimeout(() => {
      setStageIndex(offset + 1);
      if (offset === PREPARATION_STAGES.length - 2) setPreviewState("ready");
    }, 160 * (offset + 1)));
    return () => timers.forEach(window.clearTimeout);
  }, [open, document, target, excelMode, numberDisplay]);

  const moveSection = (delta: number) => setSectionIndex(current => Math.min(Math.max(0, sectionCount - 1), Math.max(0, current + delta)));
  const changeZoom = (delta: number) => setZoom(current => Math.min(160, Math.max(70, current + delta)));

  const shareOutput = async () => {
    setShareState("working");
    try {
      const blob = target === "Word" ? await createWordBlob(document, numberDisplay) : createExcelBlob(document, numberDisplay, excelMode);
      const extension = target === "Word" ? "docx" : "xlsx";
      const mime = target === "Word" ? "application/vnd.openxmlformats-officedocument.wordprocessingml.document" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
      const file = new File([blob], `Nawa_${document.fileName.replace(/\.[^.]+$/, "")}.${extension}`, { type: mime });
      const browser = navigator as Navigator & { share?: (data: ShareData) => Promise<void>; canShare?: (data?: ShareData) => boolean };
      if (browser.share && (!browser.canShare || browser.canShare({ files: [file] }))) {
        await browser.share({ title: `Nawa OCR — ${target}`, text: "ملف محول من Nawa OCR", files: [file] });
        setShareState("done");
      } else {
        setShareState("error");
      }
    } catch {
      setShareState("error");
    }
  };

  const openExternalShare = (channel: "email" | "whatsapp") => {
    const subject = encodeURIComponent(`ملف Nawa OCR المحول — ${document.fileName}`);
    const body = encodeURIComponent("أرفق ملف DOCX أو XLSX الذي تم تنزيله من نافذة المعاينة.");
    const url = channel === "email" ? `mailto:?subject=${subject}&body=${body}` : `https://wa.me/?text=${body}`;
    window.open(url, "_blank", "noopener,noreferrer");
  };

  return (
    <section className="rounded-2xl border border-indigo-100 bg-indigo-50/40 p-3" aria-label={`معاينة ملف ${target}`} aria-busy={open && previewState === "loading"}>
      <div className="flex items-center justify-between gap-3">
        <div className="flex items-center gap-2"><span className="flex size-8 items-center justify-center rounded-lg bg-white text-indigo-600 shadow-sm">{target === "Excel" ? <FileSpreadsheet className="size-4" /> : <FileText className="size-4" />}</span><div><p className="text-sm font-semibold text-slate-800">معاينة ملف {target === "Excel" ? "XLSX" : "DOCX"}</p><p className="text-[11px] text-slate-500">المعاينة مطابقة للخطة قبل إنشاء ملف التنزيل.</p></div></div>
        <Button type="button" size="sm" variant="outline" className="bg-white" onClick={() => setOpen(current => !current)} aria-expanded={open}>{open ? <EyeOff className="ml-1 size-3" /> : <Eye className="ml-1 size-3" />}{open ? "إخفاء المعاينة" : "فتح المعاينة"}</Button>
      </div>

      {open && <div className="mt-3 overflow-hidden rounded-xl border border-slate-200 bg-white shadow-inner transition-[opacity,transform] duration-200 ease-out motion-reduce:transition-none" aria-live="polite">
        {previewState === "loading" ? <div className="space-y-4 p-5">
          <div className="flex items-center gap-3"><LoaderCircle className="size-6 animate-spin text-indigo-500 motion-reduce:animate-none" aria-hidden="true" /><div><p className="text-sm font-semibold text-slate-700">جارٍ تجهيز المعاينة…</p><p className="text-xs text-slate-500">المرحلة {stageIndex + 1} من {PREPARATION_STAGES.length}: {PREPARATION_STAGES[stageIndex]}</p></div></div>
          <div className="h-2 overflow-hidden rounded-full bg-slate-100" role="progressbar" aria-valuemin={0} aria-valuemax={100} aria-valuenow={Math.round(((stageIndex + 1) / PREPARATION_STAGES.length) * 100)}><div className="h-full rounded-full bg-indigo-500 transition-[width] duration-200 ease-out motion-reduce:transition-none" style={{ width: `${((stageIndex + 1) / PREPARATION_STAGES.length) * 100}%` }} /></div>
          <div className="grid grid-cols-2 gap-2 text-[11px] text-slate-500 sm:grid-cols-4">{PREPARATION_STAGES.map((stage, index) => <span key={stage} className={index <= stageIndex ? "font-semibold text-indigo-600" : ""}>{index + 1}. {stage}</span>)}</div>
        </div> : <>
          <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-100 bg-slate-50/80 px-3 py-2">
            <div className="flex items-center gap-1"><Button type="button" size="icon" variant="outline" className="size-8 bg-white" onClick={() => changeZoom(-10)} aria-label="تصغير"><Minus className="size-3" /></Button><span className="min-w-12 text-center text-xs font-semibold text-slate-600">{zoom}%</span><Button type="button" size="icon" variant="outline" className="size-8 bg-white" onClick={() => changeZoom(10)} aria-label="تكبير"><Plus className="size-3" /></Button><Button type="button" size="icon" variant="outline" className="size-8 bg-white" onClick={() => setZoom(100)} aria-label="إعادة ضبط التكبير"><RotateCcw className="size-3" /></Button></div>
            <div className="flex items-center gap-1"><Button type="button" size="icon" variant="outline" className="size-8 bg-white" onClick={() => moveSection(-1)} disabled={sectionIndex <= 0} aria-label="السابق"><ChevronRight className="size-3" /></Button><span className="min-w-24 text-center text-xs text-slate-600">{target === "Word" ? `صفحة ${Math.min(sectionIndex + 1, Math.max(1, sectionCount))} من ${Math.max(1, sectionCount)}` : `ورقة ${Math.min(sectionIndex + 1, Math.max(1, sectionCount))} من ${Math.max(1, sectionCount)}`}</span><Button type="button" size="icon" variant="outline" className="size-8 bg-white" onClick={() => moveSection(1)} disabled={sectionIndex >= sectionCount - 1} aria-label="التالي"><ChevronLeft className="size-3" /></Button></div>
          </div>
          <div className="flex flex-wrap items-center gap-2 border-b border-slate-100 px-3 py-2 text-[11px]">
            <Button type="button" size="sm" variant="outline" className="bg-white" onClick={shareOutput} disabled={shareState === "working"}><Share2 className="ml-1 size-3" />{shareState === "working" ? "جارٍ تجهيز المشاركة…" : shareState === "done" ? "تمت المشاركة" : "مشاركة الملف"}</Button>
            <Button type="button" size="sm" variant="outline" className="bg-white" onClick={() => openExternalShare("email")}><Mail className="ml-1 size-3" />البريد</Button>
            <Button type="button" size="sm" variant="outline" className="bg-white" onClick={() => openExternalShare("whatsapp")}><Copy className="ml-1 size-3" />واتساب</Button>
            {shareState === "error" && <span className="text-amber-700">المشاركة المباشرة غير متاحة؛ نزّل الملف ثم أرفقه يدويًا.</span>}
            <span className="mr-auto text-slate-400">لا يتم رفع الملف تلقائيًا</span>
          </div>
          <div className="max-h-[28rem] overflow-auto p-4 animate-in fade-in-0 slide-in-from-bottom-1 duration-200 motion-reduce:animate-none" style={{ fontSize: `${zoom}%` }}>
            {target === "Excel" && currentWorksheet ? <section className="space-y-5"><div className="mb-2 flex items-center justify-between border-b border-slate-100 pb-2"><h3 className="font-semibold text-slate-800">{currentWorksheet.name}</h3><span className="text-[11px] text-slate-400">{currentWorksheet.tables.length} جدول</span></div>{currentWorksheet.tables.map(table => <div key={`${currentWorksheet.name}-${table.name}`} className="mb-4 overflow-auto rounded-lg border border-slate-200"><p className="bg-slate-50 px-3 py-2 text-xs font-semibold text-slate-600">{table.name} · {table.rowCount} × {table.columnCount}</p><table className="min-w-full text-right text-sm"><tbody>{table.values.map((row, rowIndex) => <tr key={rowIndex} className="border-b last:border-0">{row.map((value, columnIndex) => <td key={columnIndex} className={`whitespace-nowrap px-3 py-2 ${rowIndex === 0 ? "bg-indigo-50/60 font-semibold" : "text-slate-700"}`} dir="auto">{value == null ? "" : typeof value === "number" ? String(value) : formatDigits(String(value), numberDisplay)}</td>)}</tr>)}</tbody></table></div>)}</section> : target === "Word" && currentPage ? <section className="space-y-5" dir="auto"><p className="mb-2 text-[11px] font-semibold uppercase tracking-wide text-indigo-500">صفحة {currentPage.pageNumber}</p>{[...currentPage.header, ...currentPage.paragraphs, ...currentPage.footer].map((paragraph, index) => <p key={`${currentPage.pageNumber}-${index}`} dir={paragraph.direction} className={`mb-2 whitespace-pre-wrap leading-7 ${paragraph.kind === "heading" ? "text-base font-bold text-slate-900" : "text-sm text-slate-700"} ${paragraph.direction === "rtl" ? "text-right" : "text-left"}`}>{formatDigits(paragraph.text, numberDisplay)}</p>)}{currentPage.tables.map((table, tableIndex) => <table key={`${currentPage.pageNumber}-table-${tableIndex}`} className="mb-3 min-w-full border-collapse text-sm"><tbody>{table.rows.map((row, rowIndex) => <tr key={rowIndex}>{row.map((cell, cellIndex) => <td key={cellIndex} className={`border border-slate-200 px-3 py-2 ${rowIndex === 0 ? "bg-indigo-50/60 font-semibold" : ""}`} dir="auto">{formatDigits(cell, numberDisplay)}</td>)}</tr>)}</tbody></table>)}</section> : <p className="p-8 text-center text-sm text-slate-500">لا توجد صفحات أو أوراق للعرض.</p>}
          </div>
        </>}
      </div>}
    </section>
  );
}

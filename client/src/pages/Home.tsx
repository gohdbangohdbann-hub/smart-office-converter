/// <reference types="office-js" />
import { useEffect, useRef, useState } from "react";
import { FileUp, FileText, Files, Grid3X3, Languages, Loader2, Settings2, ShieldCheck, Sparkles, WandSparkles } from "lucide-react";
import { toast } from "sonner";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Progress } from "@/components/ui/progress";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { trpc } from "@/lib/trpc";
import { toPreviewModel, type OCRDocument } from "@shared/ocr";
import { buildWordPlan, type WordDocumentPlan } from "@shared/word";
import { buildExcelPlan, resolveExcelOperation, type ExcelWorkbookPlan } from "@shared/excel";
import { progressFor, type WordProgress } from "@shared/word-progress";
import { reduceWordProgress } from "@shared/word-progress-flow";
import { submitExcelTransform } from "@shared/excel-request";
import { uiCopy, uiDirection, type UiLocale } from "@shared/i18n";
import { verifyDocument } from "@shared/verification";
import { type BatchOutputMode, type BatchStatus } from "@shared/batch";
import { type NawaSettings } from "@shared/settings";
import { BatchPanel } from "@/components/BatchPanel";
import { SettingsPanel, loadNawaSettings } from "@/components/SettingsPanel";
import { downloadExcel, downloadExcelBatch, downloadWord, downloadWordBatch } from "@/lib/export";
import { VerificationPanel } from "@/components/VerificationPanel";
import { OutputPreview } from "@/components/OutputPreview";
import { hasOccupiedCells, tableStart, type ExcelInsertMode, type WordInsertMode } from "@shared/office-insertion";

const supported = ".pdf,.png,.jpg,.jpeg,.tiff,.webp";

type Host = "Word" | "Excel" | "Browser";


function currentHost(): Host {
  if (typeof Office === "undefined" || !Office.context?.host) return "Browser";
  if (Office.context.host === Office.HostType.Word) return "Word";
  if (Office.context.host === Office.HostType.Excel) return "Excel";
  return "Browser";
}

function officeErrorMessage(error: unknown) {
  const code = typeof error === "object" && error && "code" in error ? String((error as { code?: unknown }).code) : "";
  if (code.includes("AccessDenied") || code.includes("Permission")) return "لا يملك Office صلاحية الإدراج في هذا المستند.";
  if (code.includes("InvalidArgument")) return "تعذر إدراج النتيجة بسبب تنسيق غير صالح.";
  if (code.includes("Network")) return "تعذر الاتصال بواجهة Office. أعد فتح Task Pane وحاول مجددًا.";
  if (code.includes("EXCEL_TARGET_NOT_EMPTY")) return "منطقة الإدراج في Excel غير فارغة؛ لم تُحذف أي بيانات. اختر منطقة أخرى أو أفرغها يدويًا ثم أعد المحاولة.";
  return "تعذر الإدراج في Office. تأكد من فتح المستند وتفعيل الإضافة.";
}

async function insertIntoOffice(document: OCRDocument, host: Host, excelMode: "separate" | "single" | "smart", wordInsertMode: WordInsertMode = "cursor", excelInsertMode: ExcelInsertMode = "active-cell") {
  if (host === "Word" && typeof Word !== "undefined") {
    const plan = buildWordPlan(document);
    await Word.run(async (context: Word.RequestContext) => {
      const body = context.document.body;
      const selection = context.document.getSelection();
      let firstInsertion = true;
      for (const page of plan.pages) {
        for (const paragraph of [...page.header, ...page.paragraphs, ...page.footer]) {
          if (wordInsertMode === "replace-selection" && firstInsertion) {
            selection.insertText(paragraph.text, Word.InsertLocation.replace);
            firstInsertion = false;
            continue;
          }
          const item = wordInsertMode === "document-end" ? body.insertParagraph(paragraph.text, Word.InsertLocation.end) : selection.insertParagraph(paragraph.text, Word.InsertLocation.after);
          item.alignment = paragraph.direction === "rtl" ? Word.Alignment.right : Word.Alignment.left;
          item.font.bold = Boolean(paragraph.bold);
          if (paragraph.kind === "heading") item.styleBuiltIn = Word.BuiltInStyleName.heading1;
          else if (paragraph.kind === "list") item.styleBuiltIn = Word.BuiltInStyleName.listParagraph;
          firstInsertion = false;
        }
        for (const asset of page.assets) {
          if (asset.kind === "separator") {
            const separator = wordInsertMode === "document-end" ? body.insertParagraph("────────────────", Word.InsertLocation.end) : selection.insertParagraph("────────────────", Word.InsertLocation.after);
            separator.alignment = Word.Alignment.centered;
          }
        }
        for (const table of page.tables) {
          if (!table.rows.length || !table.rows[0]?.length) continue;
          const wordTable = wordInsertMode === "document-end" ? body.insertTable(table.rows.length, table.rows[0].length, Word.InsertLocation.end, table.rows) : selection.insertTable(table.rows.length, table.rows[0].length, Word.InsertLocation.after, table.rows);
          wordTable.styleBuiltIn = Word.BuiltInStyleName.gridTable5Dark_Accent2;
          wordTable.alignment = table.direction === "rtl" ? Word.Alignment.right : Word.Alignment.left;
        }
      }
      await context.sync();
    });
    return;
  }
  if (host === "Excel" && typeof Excel !== "undefined") {
    const plan = buildExcelPlan(document, excelMode);
    await Excel.run(async (context: Excel.RequestContext) => {
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      const activeCell = context.workbook.getSelectedRange();
      activeCell.load(["rowIndex", "columnIndex"]);
      await context.sync();
      let rowOffset = 0;
      for (const worksheetPlan of plan.worksheets) {
        const sheet = excelInsertMode === "new-sheet" ? context.workbook.worksheets.add(worksheetPlan.name) : activeSheet;
        for (const table of worksheetPlan.tables) {
          if (!table.values.length || !table.values[0]?.length) continue;
          const start = tableStart(activeCell.rowIndex, activeCell.columnIndex, rowOffset, excelInsertMode);
          const range = sheet.getRangeByIndexes(start.row, start.column, table.rowCount, table.columnCount);
          range.load("values");
          await context.sync();
          const occupied = hasOccupiedCells(range.values as unknown[][]);
          if (occupied) throw new Error("EXCEL_TARGET_NOT_EMPTY");
          range.values = table.values;
          range.format.wrapText = true;
          range.format.autofitColumns();
          range.format.autofitRows();
          range.format.horizontalAlignment = table.direction === "rtl" ? "Right" : "Left";
          for (const borderIndex of ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"] as const) { const border = range.format.borders.getItem(borderIndex); border.style = "Continuous"; border.weight = "Thin"; border.color = "#d7dce5"; }
          range.getRow(0).format.font.bold = true;
          rowOffset += table.rowCount + 1;
        }
      }
      await context.sync();
    });
    return;
  }
  throw new Error("OFFICE_HOST_UNAVAILABLE");
}

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [result, setResult] = useState<OCRDocument | null>(null);
  const [host, setHost] = useState<Host>(currentHost);
  const [language, setLanguage] = useState<string>(() => loadNawaSettings().ocrLanguage);
  const [uiLocale, setUiLocale] = useState<UiLocale>(() => loadNawaSettings().uiLocale);
  const [settings, setSettings] = useState<NawaSettings>(() => loadNawaSettings());
  const [ribbonPanel] = useState(() => typeof window === "undefined" ? "" : new URLSearchParams(window.location.search).get("panel") ?? "");
  const [showSettings, setShowSettings] = useState(() => ribbonPanel === "settings");
  const [showBatch, setShowBatch] = useState(false);
  const [target, setTarget] = useState<"Word" | "Excel">(() => currentHost() === "Excel" ? "Excel" : "Word");
  const [targetMode, setTargetMode] = useState<"auto" | "manual">("auto");
  const batchDocuments = useRef<OCRDocument[]>([]);
  const copy = uiCopy[uiLocale];
  const [wordPlan, setWordPlan] = useState<WordDocumentPlan | null>(null);
  const [excelPlan, setExcelPlan] = useState<ExcelWorkbookPlan | null>(null);
  const [metadata, setMetadata] = useState<{ pageCount: number; sourceKind: string; detectedLanguage: string; tableCount: number } | null>(null);
  const [operation, setOperation] = useState<"smart" | "pdf" | "image">("smart");
  const [excelMode, setExcelMode] = useState<"separate" | "single" | "smart">("smart");
  const [wordInsertMode, setWordInsertMode] = useState<WordInsertMode>("cursor");
  const [excelInsertMode, setExcelInsertMode] = useState<ExcelInsertMode>("active-cell");
  const [wordProgress, setWordProgress] = useState<WordProgress>(() => progressFor("idle"));
  const [status, setStatus] = useState("جاهز لاستقبال ملف");
  const inspect = trpc.word.inspect.useMutation({ onSuccess: (data) => { setMetadata(data); setWordProgress((state) => reduceWordProgress(state, { type: "file-selected", totalPages: data.pageCount })); }, onError: (error) => toast.error(error.message) });
  const process = trpc.word.transform.useMutation({ onSuccess: ({ document, plan }) => { setResult(document); setWordPlan(plan); setExcelPlan(null); setWordProgress((state) => reduceWordProgress(state, { type: "success", partial: plan.warnings.length > 0 })); setStatus(plan.warnings.length ? copy.partial : copy.doneWord); }, onError: (error) => { setWordProgress((state) => reduceWordProgress(state, { type: "failure" })); setStatus(copy.failed); toast.error(error.message); } });
  const excelProcess = trpc.excel.transform.useMutation({ onSuccess: ({ document, plan }) => { setResult(document); setExcelPlan(plan); setWordPlan(null); setWordProgress((state) => reduceWordProgress(state, { type: "success", partial: plan.warnings.length > 0 })); setStatus(plan.warnings.length ? copy.partial : copy.doneExcel); }, onError: (error) => { setWordProgress((state) => reduceWordProgress(state, { type: "failure" })); setStatus(copy.failed); toast.error(error.message); } });
  const isConverting = process.isPending || excelProcess.isPending;
  const conversionTarget: "Word" | "Excel" = host === "Word" || host === "Excel" ? host : target;
  const smartSuggestion: "Word" | "Excel" | null = metadata ? metadata.tableCount > 0 ? "Excel" : "Word" : null;
  const processBatchFile = async (batchFile: File, _taskId: string, done: (status: BatchStatus, error?: string) => void, outputMode: BatchOutputMode) => {
    try {
      const bytesBase64 = await encodeFile(batchFile);
      const payload = { fileName: batchFile.name, mimeType: batchFile.type || "application/octet-stream", bytesBase64 };
      const response = conversionTarget === "Excel" ? await excelProcess.mutateAsync({ ...payload, operation: "smart", mode: settings.excelTableMode }) : await process.mutateAsync(payload);
      const needsReview = verifyDocument(response.document).issues.length > 0;
      setFile(batchFile);
      batchDocuments.current.push(response.document);
      if (outputMode === "separate") {
        if (conversionTarget === "Excel") downloadExcel(response.document, settings.numberDisplay, settings.excelTableMode);
        else await downloadWord(response.document, settings.numberDisplay);
      }
      done(needsReview ? "review" : "completed");
    } catch (error) { done("failed", error instanceof Error ? error.message : "تعذر تحويل الملف"); }
  };

  useEffect(() => {
    if (typeof Office !== "undefined") Office.onReady(() => { const nextHost = currentHost(); setHost(nextHost); if (nextHost === "Word" || nextHost === "Excel") setTarget(nextHost); });
  }, []);
  useEffect(() => { if (uiLocale !== settings.uiLocale) setSettings(current => { const next = { ...current, uiLocale }; window.localStorage.setItem("nawa-ocr-settings-v1", JSON.stringify(next)); return next; }); }, [uiLocale, settings.uiLocale]);
  useEffect(() => {
    if (!process.isPending && !excelProcess.isPending) return;
    const stages = ["analyzing", "ocr", "building"] as const;
    let index = 0;
    const timer = window.setInterval(() => { const stage = stages[Math.min(index++, stages.length - 1)]!; const total = metadata?.pageCount ?? 1; setWordProgress((state) => reduceWordProgress(state, { type: "stage", stage, currentPage: stage === "analyzing" ? 0 : Math.min(total, index) })); setStatus(progressFor(stage, stage === "analyzing" ? 0 : Math.min(total, index), total).message); }, 700);
    return () => window.clearInterval(timer);
  }, [process.isPending, excelProcess.isPending, metadata?.pageCount]);

  const preview = result ? toPreviewModel(result) : null;
  const verification = result ? verifyDocument(result) : null;
  const confidence = Math.round((preview?.confidence ?? 0) * 100);
  const encodeFile = async (selected: File) => btoa(String.fromCharCode(...Array.from(new Uint8Array(await selected.arrayBuffer()))));
  const handleFileChange = async (selected: File | null) => {
    setFile(selected); setResult(null); setWordPlan(null); setExcelPlan(null); setMetadata(null);
    if (selected) inspect.mutate({ fileName: selected.name, mimeType: selected.type || "application/octet-stream", bytesBase64: await encodeFile(selected) });
  };
  const downloadResult = async () => {
    if (!result) return;
    const name = conversionTarget === "Excel" ? downloadExcel(result, settings.numberDisplay, settings.excelTableMode) : await downloadWord(result, settings.numberDisplay);
    toast.success(`تم تنزيل ${name}`);
  };
  const completeBatch = async (outputMode: BatchOutputMode) => {
    if (!batchDocuments.current.length || outputMode === "separate") return;
    if (conversionTarget === "Excel") downloadExcelBatch(batchDocuments.current, settings.numberDisplay, outputMode, settings.excelTableMode);
    else await downloadWordBatch(batchDocuments.current, settings.numberDisplay, outputMode);
    toast.success("تم حفظ مخرجات الدفعة");
    batchDocuments.current = [];
  };
  const openBatchResult = (review: boolean) => {
    document.querySelector("[data-results]")?.scrollIntoView({ behavior: "smooth", block: "start" });
    toast.info(review ? "تم فتح نتيجة تحتاج إلى مراجعة" : "تم فتح آخر نتيجة محولة");
  };
  const handleProcess = async () => {
    if (!file) return toast.error("اختر ملفًا أولًا");
    const base64 = await encodeFile(file);
    const initialStage = "analyzing" as const;
    const destination = operation === "smart" && targetMode === "auto" && smartSuggestion ? smartSuggestion : conversionTarget;
    if (host === "Browser" && operation === "smart" && targetMode === "auto" && smartSuggestion) setTarget(smartSuggestion);
    setWordProgress((state) => reduceWordProgress(state, { type: "start" })); setStatus(progressFor(initialStage).message);
    if (destination === "Excel") submitExcelTransform(excelProcess.mutate, file, base64, operation, excelMode);
    else process.mutate({ fileName: file.name, mimeType: file.type || "application/octet-stream", bytesBase64: base64 });
  };

  return <main dir={uiDirection(uiLocale)} className="min-h-screen overflow-x-hidden bg-[#f6f7fb] text-slate-900">
    <div className="mx-auto flex min-h-screen max-w-6xl flex-col px-5 py-6 lg:px-8">
      <header className="flex items-center justify-between border-b border-slate-200 pb-5">
        <div className="flex items-center gap-3"><div className="flex size-11 items-center justify-center rounded-2xl bg-slate-950 text-white shadow-lg"><Sparkles className="size-5" /></div><div><p className="text-xs font-semibold uppercase tracking-[0.24em] text-indigo-600">NAWA OCR OFFICE</p><h1 className="text-xl font-bold tracking-tight">{copy.title}</h1></div></div>
        <div className="flex items-center gap-2"><Button type="button" variant="outline" size="sm" className="bg-white" onClick={() => setShowSettings(current => !current)}><Settings2 className="ml-1 size-3" /> الإعدادات</Button><div className="flex rounded-lg border border-slate-200 bg-white p-0.5 text-[11px]">{([['ar','العربية'],['fr','FR'],['en','EN']] as const).map(([code, label]) => <button key={code} type="button" onClick={() => setUiLocale(code)} className={`rounded-md px-2 py-1 transition ${uiLocale === code ? "bg-slate-950 text-white" : "text-slate-500 hover:text-slate-900"}`}>{label}</button>)}</div><Badge variant="outline" className="border-slate-300 bg-white text-slate-600">Host: {host}</Badge><Badge className="bg-emerald-100 text-emerald-700 hover:bg-emerald-100"><ShieldCheck className="ml-1 size-3" /> معالجة مؤقتة</Badge></div>
      </header>

      {showSettings && <section className="pb-5"><SettingsPanel value={settings} onChange={(next) => { setSettings(next); setUiLocale(next.uiLocale); setLanguage(next.ocrLanguage); setExcelMode(next.excelTableMode); }} onClose={() => setShowSettings(false)} /></section>}
      {showBatch && <section className="pb-5"><BatchPanel destination={conversionTarget} onProcess={processBatchFile} onBatchComplete={completeBatch} onOpenResult={openBatchResult} /></section>}
      <section className="grid min-w-0 flex-1 gap-6 py-7 lg:grid-cols-[0.92fr_1.08fr]">
        <Card className="min-w-0 border-0 bg-white/90 shadow-[0_20px_60px_rgba(15,23,42,0.08)]"><CardHeader><div className="flex items-center justify-between"><div><CardTitle className="text-lg">{copy.start}</CardTitle><p className="mt-1 text-sm text-slate-500">PDF أو صورة — العربية أولًا</p></div><WandSparkles className="size-5 text-indigo-500" /></div></CardHeader><CardContent className="space-y-5">
          <label className="group flex min-h-48 cursor-pointer flex-col items-center justify-center rounded-3xl border border-dashed border-indigo-200 bg-indigo-50/40 px-5 text-center transition hover:border-indigo-400 hover:bg-indigo-50"><input type="file" accept={supported} className="sr-only" onChange={(event) => { void handleFileChange(event.target.files?.[0] ?? null); }} /><FileUp className="mb-3 size-8 text-indigo-500 transition group-hover:-translate-y-1" /><span className="font-semibold text-slate-800">{file ? file.name : copy.upload}</span><span className="mt-2 text-xs text-slate-500">PDF · PNG · JPG · TIFF · WEBP</span>{file && <span className="mt-3 text-xs text-indigo-600">{file.type || "نوع غير معروف"} · {(file.size / 1024).toFixed(1)} KB</span>}</label>
          {smartSuggestion && <div className="mb-2 rounded-lg border border-indigo-100 bg-indigo-50/60 px-3 py-2 text-xs text-indigo-800">التحويل الذكي يقترح <strong>{smartSuggestion}</strong> بناءً على التحليل الأولي؛ يمكنك تغيير الوجهة يدويًا.</div>}<div className="mb-2 grid grid-cols-2 gap-2"><Button type="button" variant="outline" className={target === "Word" ? "border-indigo-300 bg-indigo-50 text-indigo-700" : "bg-white"} onClick={() => { setTarget("Word"); setTargetMode("manual"); }}>Word</Button><Button type="button" variant="outline" className={target === "Excel" ? "border-indigo-300 bg-indigo-50 text-indigo-700" : "bg-white"} onClick={() => { setTarget("Excel"); setTargetMode("manual"); }}>Excel</Button></div><div className="grid grid-cols-3 gap-2"><Button type="button" variant="outline" className={operation === "pdf" ? "border-indigo-300 bg-indigo-50 text-indigo-700" : "bg-white"} onClick={() => setOperation("pdf")}>{conversionTarget === "Excel" ? copy.pdfExcel : copy.pdfWord}</Button><Button type="button" variant="outline" className={operation === "image" ? "border-indigo-300 bg-indigo-50 text-indigo-700" : "bg-white"} onClick={() => setOperation("image")}>{conversionTarget === "Excel" ? copy.imageExcel : copy.imageWord}</Button><Button type="button" variant="outline" className={operation === "smart" ? "border-indigo-300 bg-indigo-50 text-indigo-700" : "bg-white"} onClick={() => { setOperation("smart"); setTargetMode("auto"); }}>{conversionTarget === "Excel" ? copy.smartExcel : copy.smartWord}</Button></div>{conversionTarget === "Excel" && <div className="rounded-xl bg-slate-50 p-2"><p className="mb-2 text-xs font-medium text-slate-600">توزيع الجداول</p><div className="grid grid-cols-3 gap-1"><Button type="button" variant="outline" className={excelMode === "separate" ? "bg-indigo-50 text-indigo-700" : "bg-white"} onClick={() => setExcelMode("separate")}>{copy.sheetPerTable}</Button><Button type="button" variant="outline" className={excelMode === "single" ? "bg-indigo-50 text-indigo-700" : "bg-white"} onClick={() => setExcelMode("single")}>{copy.oneSheet}</Button><Button type="button" variant="outline" className={excelMode === "smart" ? "bg-indigo-50 text-indigo-700" : "bg-white"} onClick={() => setExcelMode("smart")}>{copy.smart}</Button></div></div>}{conversionTarget === "Word" ? <div className="grid gap-2 rounded-xl bg-slate-50 p-2"><label className="text-xs font-medium text-slate-600">مكان إدراج Word</label><select value={wordInsertMode} onChange={event => setWordInsertMode(event.target.value as WordInsertMode)} className="h-9 rounded-md border border-slate-200 bg-white px-2 text-sm"><option value="cursor">عند موضع المؤشر</option><option value="document-end">في نهاية المستند</option><option value="replace-selection">استبدال التحديد</option></select></div> : <div className="grid gap-2 rounded-xl bg-slate-50 p-2"><label className="text-xs font-medium text-slate-600">مكان إدراج Excel</label><select value={excelInsertMode} onChange={event => setExcelInsertMode(event.target.value as ExcelInsertMode)} className="h-9 rounded-md border border-slate-200 bg-white px-2 text-sm"><option value="active-cell">الخلية الحالية</option><option value="new-sheet">ورقة جديدة</option></select><p className="text-[11px] text-slate-500">لا تُستبدل البيانات الموجودة تلقائيًا؛ سيظهر تحذير عند امتلاء المنطقة.</p></div>}<div><div className="mb-2 flex items-center justify-between text-sm"><span className="font-medium">{copy.language}</span><span className="text-slate-500">{language === "auto" ? "تلقائي" : language}</span></div><Tabs value={language} onValueChange={setLanguage}><TabsList className="grid w-full grid-cols-4 bg-slate-100"><TabsTrigger value="auto">تلقائي</TabsTrigger><TabsTrigger value="ar">العربية</TabsTrigger><TabsTrigger value="fr">Français</TabsTrigger><TabsTrigger value="en">English</TabsTrigger></TabsList></Tabs></div>
          <Button type="button" variant="outline" className="h-10 w-full" onClick={() => setShowBatch(current => !current)}><Files className="ml-2 size-4" /> تحويل عدة ملفات</Button><Button className="h-12 w-full rounded-xl bg-slate-950 text-base hover:bg-indigo-700" onClick={handleProcess} disabled={!file || isConverting}>{isConverting ? <><Loader2 className="ml-2 size-4 animate-spin" /> {conversionTarget === "Excel" ? copy.processingExcel : copy.processingWord}</> : <><WandSparkles className="ml-2 size-4" /> {conversionTarget === "Excel" ? copy.smartExcel : copy.smartWord}</>}</Button>
          <div className="rounded-xl border border-slate-100 bg-slate-50 px-3 py-2 text-sm text-slate-600">{status}{isConverting && <><Progress value={wordProgress.percent} className="mt-2 h-1.5" /><span className="mt-1 block text-xs text-indigo-600">{wordProgress.percent}% · الصفحة {wordProgress.currentPage || 1} من {wordProgress.totalPages || metadata?.pageCount || 1}</span></>}{metadata && <span className="mr-2 text-xs text-slate-400">· {metadata.pageCount} صفحة · {metadata.tableCount} جدول مبدئي · {metadata.sourceKind} · اللغة المكتشفة: {metadata.detectedLanguage}</span>}{result && <span className="mr-2 text-xs text-slate-400">· اللغة النهائية: {result.language}</span>}</div><p className="flex items-center gap-2 text-xs leading-5 text-slate-500"><ShieldCheck className="size-4 shrink-0 text-emerald-600" /> {copy.privacy}</p>
        </CardContent></Card>

        <Card data-results className="min-w-0 border-0 bg-white/90 shadow-[0_20px_60px_rgba(15,23,42,0.08)]"><CardHeader><div className="flex items-center justify-between"><div><CardTitle className="text-lg">{copy.preview}</CardTitle><p className="mt-1 text-sm text-slate-500">{conversionTarget === "Excel" ? "خلايا Excel قابلة للتحرير" : "خطة Word قابلة للتحرير"}</p></div>{result && <Badge className="bg-indigo-100 text-indigo-700 hover:bg-indigo-100">{result.provider}</Badge>}</div></CardHeader><CardContent>{!result ? <div className="flex min-h-[340px] flex-col items-center justify-center rounded-2xl bg-slate-50 text-center"><FileText className="mb-4 size-10 text-slate-300" /><p className="font-medium text-slate-600">{copy.emptyResult}</p><p className="mt-1 max-w-xs text-sm leading-6 text-slate-400">{copy.emptyHint}</p></div> : <div className="space-y-4"><div className="grid grid-cols-3 gap-3"><Stat icon={<Languages />} label={copy.textLanguage} value={preview?.language ?? "—"} /><Stat icon={<FileText />} label={copy.pages} value={String(preview?.pageCount ?? 0)} /><Stat icon={<Grid3X3 />} label={copy.tables} value={String(preview?.tables.length ?? 0)} /></div><div className="rounded-2xl border border-slate-200 bg-slate-50 p-4"><div className="mb-3 flex items-center justify-between"><span className="text-sm font-semibold">{copy.confidence}</span><span className="text-sm font-bold text-indigo-700">{confidence}%</span></div><Progress value={confidence} className="h-2" /><p className="mt-2 text-xs text-slate-500">{copy.confidenceHint}</p></div>{conversionTarget === "Excel" && excelPlan ? <div className="max-h-64 overflow-auto rounded-2xl border border-slate-200 p-4 text-sm"><p className="mb-3 font-semibold">{copy.tablePreview}</p>{excelPlan.worksheets.map((worksheet) => worksheet.tables.map((table) => <div key={table.name} className="mb-4 overflow-auto rounded-lg border border-slate-200"><p className="bg-slate-50 px-3 py-2 text-xs font-semibold">{worksheet.name} · {table.rowCount} × {table.columnCount}</p><table className="min-w-full text-right"><tbody>{table.values.map((row, rowIndex) => <tr key={rowIndex} className="border-b last:border-0">{row.map((value, columnIndex) => <td key={columnIndex} className={`px-3 py-2 ${rowIndex === 0 ? "font-semibold bg-indigo-50/50" : ""}`}>{value == null ? "" : String(value)}</td>)}</tr>)}</tbody></table></div>))}</div> : <div className="max-h-52 overflow-auto rounded-2xl border border-slate-200 p-4 text-sm leading-8 whitespace-pre-wrap">{preview?.text}</div>}{verification && <VerificationPanel report={verification} sourceFile={file} />}<OutputPreview document={result} target={conversionTarget} numberDisplay={settings.numberDisplay} excelMode={settings.excelTableMode} />{wordPlan?.warnings.length ? <div className="rounded-xl border border-amber-200 bg-amber-50 p-3 text-xs text-amber-800">{wordPlan.warnings.join(" · ")}</div> : null}{(preview?.tables.length ?? 0) > 0 && <div className="rounded-2xl border border-slate-200 p-4 text-sm"><p className="mb-2 font-semibold">{copy.foundTables}</p><p className="mb-3 text-slate-500">{preview?.tables.length ?? 0} جدول · {preview?.tables[0]?.rowCount} صفوف · {preview?.tables[0]?.columnCount} أعمدة</p><div className="overflow-auto rounded-lg border border-slate-200"><table className="min-w-full text-right"><tbody>{Array.from({ length: preview?.tables[0]?.rowCount ?? 0 }, (_, row) => <tr key={row} className="border-b last:border-0">{Array.from({ length: preview?.tables[0]?.columnCount ?? 0 }, (_, column) => <td key={column} className="px-3 py-2 text-slate-700">{preview?.tables[0]?.cells.find((cell) => cell.rowIndex === row && cell.columnIndex === column)?.text ?? "—"}</td>)}</tr>)}</tbody></table></div></div>}<div className="grid gap-2 sm:grid-cols-2"><Button variant="outline" className="h-11 rounded-xl border-indigo-200 bg-indigo-50 text-indigo-700 hover:bg-indigo-100" onClick={() => insertIntoOffice(result, conversionTarget, excelMode, wordInsertMode, excelInsertMode).then(() => toast.success(conversionTarget === "Excel" ? "تم إدراج الخلايا الحقيقية في Excel" : "تم إدراج المحتوى القابل للتحرير في Word")).catch((error) => toast.error(officeErrorMessage(error)))}>{conversionTarget === "Excel" ? copy.insertExcel : copy.insertWord}</Button><Button variant="outline" className="h-11 rounded-xl" onClick={() => { void downloadResult(); }}>تنزيل {conversionTarget === "Excel" ? ".xlsx" : ".docx"}</Button></div></div>}</CardContent></Card>
      </section>
      <footer className="flex flex-wrap items-center justify-between gap-3 border-t border-slate-200 pt-4 text-xs text-slate-500"><span>المرحلة {host === "Excel" ? "الثالثة" : "الثانية"} · Nawa OCR {host}</span><span>PDF / Image → {host === "Excel" ? "Excel" : "Word"} · العربية أولًا</span></footer>
    </div>
  </main>;
}

function Stat({ icon, label, value }: { icon: React.ReactNode; label: string; value: string }) { return <div className="rounded-2xl bg-slate-50 p-3"><div className="mb-2 flex items-center gap-2 text-slate-400"><span className="[&>svg]:size-4">{icon}</span><span className="text-xs">{label}</span></div><p className="font-bold text-slate-800">{value}</p></div>; }

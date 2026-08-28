/// <reference types="office-js" />
import { useEffect, useState } from "react";
import { FileUp, FileText, Grid3X3, Languages, Loader2, ShieldCheck, Sparkles, WandSparkles } from "lucide-react";
import { toast } from "sonner";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Progress } from "@/components/ui/progress";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { trpc } from "@/lib/trpc";
import { toPreviewModel, type OCRDocument } from "@shared/ocr";

const supported = ".pdf,.png,.jpg,.jpeg,.tiff,.webp";

type Host = "Word" | "Excel" | "Browser";
type UiLocale = "ar" | "fr" | "en";

function currentHost(): Host {
  if (typeof Office === "undefined" || !Office.context?.host) return "Browser";
  return Office.context.host === Office.HostType.Word ? "Word" : Office.context.host === Office.HostType.Excel ? "Excel" : "Browser";
}

async function insertIntoOffice(document: OCRDocument, host: Host) {
  if (host === "Word" && typeof Word !== "undefined") {
    const text = toPreviewModel(document).text;
    await Word.run(async (context: Word.RequestContext) => { context.document.body.insertText(text, Word.InsertLocation.end); await context.sync(); });
    return;
  }
  if (host === "Excel" && typeof Excel !== "undefined") {
    const table = document.tables[0];
    const rows = table ? Array.from({ length: table.rowCount }, (_, row) => Array.from({ length: table.columnCount }, (_, column) => table.cells.find((cell) => cell.rowIndex === row && cell.columnIndex === column)?.text ?? "")) : [[toPreviewModel(document).text]];
    await Excel.run(async (context) => { const range = context.workbook.worksheets.getActiveWorksheet().getRangeByIndexes(0, 0, rows.length, Math.max(1, rows[0]?.length ?? 1)); range.values = rows; range.format.autofitColumns(); await context.sync(); });
    return;
  }
  throw new Error("OFFICE_HOST_UNAVAILABLE");
}

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [result, setResult] = useState<OCRDocument | null>(null);
  const [host, setHost] = useState<Host>(currentHost);
  const [language, setLanguage] = useState("auto");
  const [uiLocale, setUiLocale] = useState<UiLocale>("ar");
  const process = trpc.ocr.process.useMutation({ onSuccess: setResult, onError: (error) => toast.error(error.message) });

  useEffect(() => {
    if (typeof Office !== "undefined") Office.onReady(() => setHost(currentHost()));
  }, []);

  const preview = result ? toPreviewModel(result) : null;
  const confidence = Math.round((preview?.confidence ?? 0) * 100);
  const handleProcess = async () => {
    if (!file) return toast.error("اختر ملفًا أولًا");
    const bytes = await file.arrayBuffer();
    const base64 = btoa(String.fromCharCode(...Array.from(new Uint8Array(bytes))));
    process.mutate({ fileName: file.name, mimeType: file.type || "application/octet-stream", bytesBase64: base64 });
  };

  return <main dir={uiLocale === "ar" ? "rtl" : "ltr"} className="min-h-screen overflow-x-hidden bg-[#f6f7fb] text-slate-900">
    <div className="mx-auto flex min-h-screen max-w-6xl flex-col px-5 py-6 lg:px-8">
      <header className="flex items-center justify-between border-b border-slate-200 pb-5">
        <div className="flex items-center gap-3"><div className="flex size-11 items-center justify-center rounded-2xl bg-slate-950 text-white shadow-lg"><Sparkles className="size-5" /></div><div><p className="text-xs font-semibold uppercase tracking-[0.24em] text-indigo-600">NAWA OCR OFFICE</p><h1 className="text-xl font-bold tracking-tight">مساحة التحويل الذكي</h1></div></div>
        <div className="flex items-center gap-2"><div className="flex rounded-lg border border-slate-200 bg-white p-0.5 text-[11px]">{([['ar','العربية'],['fr','FR'],['en','EN']] as const).map(([code, label]) => <button key={code} type="button" onClick={() => setUiLocale(code)} className={`rounded-md px-2 py-1 transition ${uiLocale === code ? "bg-slate-950 text-white" : "text-slate-500 hover:text-slate-900"}`}>{label}</button>)}</div><Badge variant="outline" className="border-slate-300 bg-white text-slate-600">Host: {host}</Badge><Badge className="bg-emerald-100 text-emerald-700 hover:bg-emerald-100"><ShieldCheck className="ml-1 size-3" /> معالجة مؤقتة</Badge></div>
      </header>

      <section className="grid min-w-0 flex-1 gap-6 py-7 lg:grid-cols-[0.92fr_1.08fr]">
        <Card className="min-w-0 border-0 bg-white/90 shadow-[0_20px_60px_rgba(15,23,42,0.08)]"><CardHeader><div className="flex items-center justify-between"><div><CardTitle className="text-lg">ابدأ من ملفك</CardTitle><p className="mt-1 text-sm text-slate-500">PDF أو صورة — العربية أولًا</p></div><WandSparkles className="size-5 text-indigo-500" /></div></CardHeader><CardContent className="space-y-5">
          <label className="group flex min-h-48 cursor-pointer flex-col items-center justify-center rounded-3xl border border-dashed border-indigo-200 bg-indigo-50/40 px-5 text-center transition hover:border-indigo-400 hover:bg-indigo-50"><input type="file" accept={supported} className="sr-only" onChange={(event) => { setFile(event.target.files?.[0] ?? null); setResult(null); }} /><FileUp className="mb-3 size-8 text-indigo-500 transition group-hover:-translate-y-1" /><span className="font-semibold text-slate-800">{file ? file.name : "اسحب الملف هنا أو اختره"}</span><span className="mt-2 text-xs text-slate-500">PDF · PNG · JPG · TIFF · WEBP</span></label>
          <div><div className="mb-2 flex items-center justify-between text-sm"><span className="font-medium">اكتشاف اللغة</span><span className="text-slate-500">{language === "auto" ? "تلقائي" : language}</span></div><Tabs value={language} onValueChange={setLanguage}><TabsList className="grid w-full grid-cols-4 bg-slate-100"><TabsTrigger value="auto">تلقائي</TabsTrigger><TabsTrigger value="ar">العربية</TabsTrigger><TabsTrigger value="fr">Français</TabsTrigger><TabsTrigger value="en">English</TabsTrigger></TabsList></Tabs></div>
          <Button className="h-12 w-full rounded-xl bg-slate-950 text-base hover:bg-indigo-700" onClick={handleProcess} disabled={!file || process.isPending}>{process.isPending ? <><Loader2 className="ml-2 size-4 animate-spin" /> جارٍ التحليل...</> : <><WandSparkles className="ml-2 size-4" /> تحويل ذكي</>}</Button>
          <p className="flex items-center gap-2 text-xs leading-5 text-slate-500"><ShieldCheck className="size-4 shrink-0 text-emerald-600" /> لا يُحتفظ بالملف الأصلي؛ تُمسح البيانات المؤقتة بعد اكتمال المعالجة.</p>
        </CardContent></Card>

        <Card className="min-w-0 border-0 bg-white/90 shadow-[0_20px_60px_rgba(15,23,42,0.08)]"><CardHeader><div className="flex items-center justify-between"><div><CardTitle className="text-lg">معاينة النتيجة</CardTitle><p className="mt-1 text-sm text-slate-500">نموذج موحّد مستقل عن Word وExcel</p></div>{result && <Badge className="bg-indigo-100 text-indigo-700 hover:bg-indigo-100">{result.provider}</Badge>}</div></CardHeader><CardContent>{!result ? <div className="flex min-h-[340px] flex-col items-center justify-center rounded-2xl bg-slate-50 text-center"><FileText className="mb-4 size-10 text-slate-300" /><p className="font-medium text-slate-600">ستظهر النتيجة هنا</p><p className="mt-1 max-w-xs text-sm leading-6 text-slate-400">النص، الثقة، اللغة، والجداول في بنية واحدة قابلة للإدراج.</p></div> : <div className="space-y-4"><div className="grid grid-cols-3 gap-3"><Stat icon={<Languages />} label="اللغة" value={preview?.language ?? "—"} /><Stat icon={<FileText />} label="الصفحات" value={String(preview?.pageCount ?? 0)} /><Stat icon={<Grid3X3 />} label="الجداول" value={String(preview?.tables.length ?? 0)} /></div><div className="rounded-2xl border border-slate-200 bg-slate-50 p-4"><div className="mb-3 flex items-center justify-between"><span className="text-sm font-semibold">الثقة الإرشادية</span><span className="text-sm font-bold text-indigo-700">{confidence}%</span></div><Progress value={confidence} className="h-2" /><p className="mt-2 text-xs text-slate-500">مؤشر للمراجعة وليس ضمانًا لصحة النص.</p></div><div className="max-h-52 overflow-auto rounded-2xl border border-slate-200 p-4 text-sm leading-8 whitespace-pre-wrap">{preview?.text}</div>{(preview?.tables.length ?? 0) > 0 && <div className="rounded-2xl border border-slate-200 p-4 text-sm"><p className="mb-2 font-semibold">الجداول المكتشفة</p><p className="mb-3 text-slate-500">{preview?.tables.length ?? 0} جدول · {preview?.tables[0]?.rowCount} صفوف · {preview?.tables[0]?.columnCount} أعمدة</p><div className="overflow-auto rounded-lg border border-slate-200"><table className="min-w-full text-right"><tbody>{Array.from({ length: preview?.tables[0]?.rowCount ?? 0 }, (_, row) => <tr key={row} className="border-b last:border-0">{Array.from({ length: preview?.tables[0]?.columnCount ?? 0 }, (_, column) => <td key={column} className="px-3 py-2 text-slate-700">{preview?.tables[0]?.cells.find((cell) => cell.rowIndex === row && cell.columnIndex === column)?.text ?? "—"}</td>)}</tr>)}</tbody></table></div></div>}<Button variant="outline" className="h-11 w-full rounded-xl border-indigo-200 bg-indigo-50 text-indigo-700 hover:bg-indigo-100" onClick={() => insertIntoOffice(result, host).then(() => toast.success(host === "Word" ? "تم إدراج النص في Word" : "تم إدراج البيانات في Excel")).catch(() => toast.error("افتح الإضافة داخل Word أو Excel لتنفيذ الإدراج"))}>إدراج في {host === "Excel" ? "Excel" : host === "Word" ? "Word" : "Office"}</Button></div>}</CardContent></Card>
      </section>
      <footer className="flex flex-wrap items-center justify-between gap-3 border-t border-slate-200 pt-4 text-xs text-slate-500"><span>المرحلة الأولى · OCR Core Preview</span><span>العربية · Français · English</span></footer>
    </div>
  </main>;
}

function Stat({ icon, label, value }: { icon: React.ReactNode; label: string; value: string }) { return <div className="rounded-2xl bg-slate-50 p-3"><div className="mb-2 flex items-center gap-2 text-slate-400"><span className="[&>svg]:size-4">{icon}</span><span className="text-xs">{label}</span></div><p className="font-bold text-slate-800">{value}</p></div>; }

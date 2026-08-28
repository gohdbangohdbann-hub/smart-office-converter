import { useEffect, useState } from "react";
import { RotateCcw, Settings2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { DEFAULT_SETTINGS, SETTINGS_STORAGE_KEY, type NawaSettings } from "@shared/settings";

export function loadNawaSettings(): NawaSettings {
  if (typeof window === "undefined") return DEFAULT_SETTINGS;
  try { const raw = window.localStorage.getItem(SETTINGS_STORAGE_KEY); return raw ? { ...DEFAULT_SETTINGS, ...JSON.parse(raw) } : DEFAULT_SETTINGS; } catch { return DEFAULT_SETTINGS; }
}

export function SettingsPanel({ value, onChange, onClose }: { value: NawaSettings; onChange: (value: NawaSettings) => void; onClose: () => void }) {
  const [draft, setDraft] = useState(value);
  useEffect(() => setDraft(value), [value]);
  const update = <K extends keyof NawaSettings>(key: K, next: NawaSettings[K]) => setDraft(current => ({ ...current, [key]: next }));
  const save = () => { window.localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(draft)); onChange(draft); onClose(); };
  const reset = () => { window.localStorage.removeItem(SETTINGS_STORAGE_KEY); setDraft(DEFAULT_SETTINGS); onChange(DEFAULT_SETTINGS); };
  const copy = draft.uiLocale === "fr" ? { title: "Paramètres", close: "Fermer", ui: "Langue de l’interface", ocr: "Langue OCR", digits: "Affichage des chiffres", quality: "Qualité Word", wordDirection: "Direction Word", font: "Police de secours", sheets: "Tableaux Excel", numeric: "Cellules numériques", save: "Enregistrer", reset: "Réinitialiser" } : draft.uiLocale === "en" ? { title: "Settings", close: "Close", ui: "Interface language", ocr: "OCR language", digits: "Number display", quality: "Word quality", wordDirection: "Word direction", font: "Fallback font", sheets: "Excel tables", numeric: "Numeric cells", save: "Save preferences", reset: "Reset defaults" } : { title: "الإعدادات", close: "إغلاق", ui: "لغة الواجهة", ocr: "لغة OCR", digits: "طريقة عرض الأرقام", quality: "جودة Word", wordDirection: "اتجاه Word", font: "الخط البديل", sheets: "جداول Excel", numeric: "الخلايا الرقمية", save: "حفظ التفضيلات", reset: "استعادة الافتراضي" };
  return <Card className="border-indigo-100 bg-white shadow-sm"><CardHeader className="flex-row items-center justify-between"><CardTitle className="flex items-center gap-2 text-base"><Settings2 className="size-4 text-indigo-600" /> {copy.title}</CardTitle><Button variant="ghost" size="sm" onClick={onClose}>{copy.close}</Button></CardHeader><CardContent className="grid gap-4 sm:grid-cols-2">
    <Field label={copy.ui}><select value={draft.uiLocale} onChange={e => update("uiLocale", e.target.value as NawaSettings["uiLocale"])}><option value="ar">العربية</option><option value="fr">Français</option><option value="en">English</option></select></Field>
    <Field label={copy.ocr}><select value={draft.ocrLanguage} onChange={e => update("ocrLanguage", e.target.value as NawaSettings["ocrLanguage"])}><option value="auto">تلقائي</option><option value="ar">العربية</option><option value="fr">Français</option><option value="en">English</option><option value="ar-fr">عربي + فرنسي</option><option value="ar-en">عربي + إنجليزي</option><option value="ar-fr-en">عربي + فرنسي + إنجليزي</option></select></Field>
    <Field label={copy.digits}><select value={draft.numberDisplay} onChange={e => update("numberDisplay", e.target.value as NawaSettings["numberDisplay"])}><option value="western">أرقام غربية 0123456789</option><option value="arabic">أرقام عربية ٠١٢٣٤٥٦٧٨٩</option></select></Field>
    <Field label={copy.quality}><select value={draft.wordQuality} onChange={e => update("wordQuality", e.target.value as NawaSettings["wordQuality"])}><option value="highest">أعلى دقة</option><option value="balanced">متوازن</option><option value="fast">أسرع</option></select></Field>
    <Field label={copy.wordDirection}><select value={draft.documentDirection} onChange={e => update("documentDirection", e.target.value as NawaSettings["documentDirection"])}><option value="auto">تلقائي</option><option value="rtl">RTL</option><option value="ltr">LTR</option></select></Field>
    <Field label={copy.font}><input value={draft.fallbackFont} onChange={e => update("fallbackFont", e.target.value)} /></Field>
    <Field label={copy.sheets}><select value={draft.excelTableMode} onChange={e => update("excelTableMode", e.target.value as NawaSettings["excelTableMode"])}><option value="smart">تلقائي</option><option value="separate">ورقة لكل جدول</option><option value="single">كل الجداول في ورقة واحدة</option></select></Field>
    <Field label={copy.numeric}><select value={draft.excelNumericMode} onChange={e => update("excelNumericMode", e.target.value as NawaSettings["excelNumericMode"])}><option value="auto">اكتشاف تلقائي</option><option value="text">إبقاء النص كما هو</option></select></Field>
    <div className="flex gap-2 sm:col-span-2"><Button onClick={save}>{copy.save}</Button><Button variant="outline" onClick={reset}><RotateCcw className="ml-1 size-3" /> {copy.reset}</Button></div>
  </CardContent></Card>;
}
function Field({ label, children }: { label: string; children: React.ReactNode }) { return <label className="space-y-1 text-xs font-medium text-slate-600">{label}<span className="block [&>input]:w-full [&>input]:rounded-md [&>input]:border [&>input]:border-slate-200 [&>input]:bg-slate-50 [&>input]:px-3 [&>input]:py-2 [&>select]:w-full [&>select]:rounded-md [&>select]:border [&>select]:border-slate-200 [&>select]:bg-slate-50 [&>select]:px-3 [&>select]:py-2">{children}</span></label>; }

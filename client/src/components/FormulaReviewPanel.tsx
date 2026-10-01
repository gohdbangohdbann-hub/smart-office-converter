import { useMemo, useState } from "react";
import { Check, Pencil, ShieldAlert, X } from "lucide-react";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { inferSafeFormulas, updateSuggestionStatus, type FormulaInferenceResult, type FormulaSuggestion } from "@shared/formula-inference";
import type { ExcelTablePlan } from "@shared/excel";

interface FormulaReviewPanelProps {
  tables: ExcelTablePlan[];
  onApproved?: (suggestions: FormulaSuggestion[]) => void;
}
export function FormulaReviewPanel({ tables, onApproved }: FormulaReviewPanelProps) {
  const results = useMemo<FormulaInferenceResult[]>(() => tables.map(inferSafeFormulas), [tables]);
  const [state, setState] = useState<Record<string, FormulaSuggestion>>({});
  const suggestions = results.flatMap((result) => result.suggestions).map((suggestion) => state[suggestion.id] ?? suggestion);
  const approved = suggestions.filter((suggestion) => suggestion.status === "accepted" || suggestion.status === "edited");
  const update = (suggestion: FormulaSuggestion, status: FormulaSuggestion["status"], formula?: string) => {
    try { setState((current) => ({ ...current, [suggestion.id]: updateSuggestionStatus(suggestion, status, formula) })); }
    catch { /* لا نسمح بتثبيت صيغة لا تبدأ بـ = */ }
  };
  return <Card dir="rtl" className="border-amber-200 bg-amber-50/40">
    <CardHeader className="pb-3"><div className="flex items-center justify-between gap-3"><div><CardTitle className="text-base">مراجعة المعادلات المقترحة</CardTitle><p className="mt-1 text-xs text-slate-600">لا تُدرج أي معادلة تلقائيًا؛ يجب قبولها أو تعديلها يدويًا أولًا.</p></div><ShieldAlert className="size-5 text-amber-600" /></div></CardHeader>
    <CardContent className="space-y-3">
      {!suggestions.length ? <p className="rounded-lg bg-white/70 p-3 text-sm text-slate-600">لم يتم العثور على علاقة رقمية آمنة لاقتراح معادلة.</p> : suggestions.map((suggestion) => <FormulaRow key={suggestion.id} suggestion={suggestion} onUpdate={update} />)}
      {results.flatMap((result) => result.warnings).map((warning) => <p key={warning} className="text-xs text-amber-800">{warning}</p>)}
      <div className="flex items-center justify-between border-t border-amber-200 pt-3"><span className="text-xs text-slate-600">المقبول للمراجعة النهائية: {approved.length}</span><Button type="button" size="sm" disabled={!approved.length} onClick={() => onApproved?.(approved)}>تأكيد قائمة المقبول</Button></div>
    </CardContent>
  </Card>;
}
function FormulaRow({ suggestion, onUpdate }: { suggestion: FormulaSuggestion; onUpdate: (suggestion: FormulaSuggestion, status: FormulaSuggestion["status"], formula?: string) => void }) {
  const [editing, setEditing] = useState(false);
  const [formula, setFormula] = useState(suggestion.formula);
  const accepted = suggestion.status === "accepted" || suggestion.status === "edited";
  return <div className={`rounded-lg border bg-white p-3 ${accepted ? "border-emerald-300" : "border-slate-200"}`}>
    <div className="flex flex-wrap items-center gap-2"><Badge variant="outline">{suggestion.address}</Badge><code dir="ltr" className="rounded bg-slate-100 px-2 py-1 text-xs">{suggestion.formula}</code><span className="text-xs text-slate-500">ثقة {Math.round(suggestion.confidence * 100)}%</span><span className="mr-auto text-xs text-slate-500">{suggestion.targetHeader}</span></div>
    <p className="mt-2 text-xs text-slate-600">{suggestion.reason}</p>
    {editing && <Input dir="ltr" value={formula} onChange={(event) => setFormula(event.target.value)} className="mt-2 bg-white text-left" aria-label={`تعديل معادلة ${suggestion.address}`} />}
    <div className="mt-3 flex flex-wrap gap-2">
      {editing ? <><Button type="button" size="sm" onClick={() => { onUpdate(suggestion, "edited", formula); setEditing(false); }}><Check className="ml-1 size-3" />حفظ التعديل</Button><Button type="button" size="sm" variant="outline" onClick={() => setEditing(false)}>إلغاء</Button></> : <><Button type="button" size="sm" onClick={() => onUpdate(suggestion, "accepted")}><Check className="ml-1 size-3" />قبول</Button><Button type="button" size="sm" variant="outline" onClick={() => setEditing(true)}><Pencil className="ml-1 size-3" />تعديل</Button><Button type="button" size="sm" variant="ghost" onClick={() => onUpdate(suggestion, "rejected")}><X className="ml-1 size-3" />رفض</Button></>}
    </div>
  </div>;
}

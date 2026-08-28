export type WordStage = "idle" | "analyzing" | "ocr" | "building" | "done" | "partial-failure" | "error";
export interface WordProgress { stage: WordStage; percent: number; currentPage: number; totalPages: number; message: string; }

export function progressFor(stage: WordStage, currentPage = 0, totalPages = 0): WordProgress {
  const messages: Record<WordStage, string> = { idle: "جاهز لاستقبال ملف", analyzing: "جاري تحليل المستند...", ocr: "جاري التعرف على النص...", building: "جاري بناء مستند Word...", done: "اكتمل بناء مستند Word", "partial-failure": "اكتمل التحويل مع صفحات تحتاج مراجعة", error: "تعذر إكمال التحويل" };
  const base: Record<WordStage, number> = { idle: 0, analyzing: 15, ocr: 45, building: 80, done: 100, "partial-failure": 100, error: 0 };
  const pagePercent = totalPages > 0 && (stage === "ocr" || stage === "building") ? Math.min(35, Math.round((currentPage / totalPages) * 35)) : 0;
  return { stage, percent: Math.min(100, base[stage] + pagePercent), currentPage, totalPages, message: messages[stage] };
}

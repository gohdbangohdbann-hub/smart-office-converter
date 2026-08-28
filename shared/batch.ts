export type BatchStatus = "queued" | "processing" | "completed" | "review" | "failed" | "cancelled";
export type BatchOutputMode = "separate" | "combined" | "file-sheets";

export interface BatchTask {
  id: string;
  fileName: string;
  mimeType: string;
  status: BatchStatus;
  outputName?: string;
  error?: string;
}

export function outputFileName(fileName: string, destination: "Word" | "Excel", suffix = ""): string {
  const base = fileName.replace(/\.[^.]+$/, "").replace(/[\\/:*?"<>|]+/g, "_").replace(/\s+/g, "_") || "document";
  return `${base}_Converted${suffix}.${destination === "Word" ? "docx" : "xlsx"}`;
}

export function nextOutputFileName(fileName: string, destination: "Word" | "Excel", existing: string[]): string {
  const first = outputFileName(fileName, destination);
  if (!existing.includes(first)) return first;
  let n = 2;
  while (existing.includes(outputFileName(fileName, destination, `_${n}`))) n += 1;
  return outputFileName(fileName, destination, `_${n}`);
}

export function createBatchTasks(files: Array<{ name: string; type: string }>): BatchTask[] {
  return files.map((file, index) => ({ id: `${Date.now()}-${index}-${file.name}`, fileName: file.name, mimeType: file.type || "application/octet-stream", status: "queued" }));
}

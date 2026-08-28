import { progressFor, type WordProgress, type WordStage } from "./word-progress";

export type WordProgressEvent =
  | { type: "file-selected"; totalPages: number }
  | { type: "start" }
  | { type: "stage"; stage: Exclude<WordStage, "idle" | "done" | "error" | "partial-failure">; currentPage?: number }
  | { type: "success"; partial?: boolean }
  | { type: "failure" };

export function reduceWordProgress(state: WordProgress, event: WordProgressEvent): WordProgress {
  if (event.type === "file-selected") return progressFor("idle", 0, event.totalPages);
  if (event.type === "start") return progressFor("analyzing", 0, state.totalPages);
  if (event.type === "stage") return progressFor(event.stage, event.currentPage ?? state.currentPage, state.totalPages);
  if (event.type === "success") return progressFor(event.partial ? "partial-failure" : "done", state.totalPages, state.totalPages);
  return progressFor("error", 0, state.totalPages);
}

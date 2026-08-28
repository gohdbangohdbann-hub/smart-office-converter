import { resolveExcelOperation } from "./excel";

export interface ExcelTransformRequest { fileName: string; mimeType: string; bytesBase64: string; operation: "pdf" | "image" | "smart"; mode: "separate" | "single" | "smart"; }
export function buildExcelTransformRequest(file: Pick<File, "name" | "type">, bytesBase64: string, operation: ExcelTransformRequest["operation"], mode: ExcelTransformRequest["mode"]): ExcelTransformRequest { return { fileName: file.name, mimeType: file.type || "application/octet-stream", bytesBase64, operation: resolveExcelOperation(operation, file.name, file.type || "application/octet-stream"), mode }; }
export function submitExcelTransform(mutate: (request: ExcelTransformRequest) => unknown, file: Pick<File, "name" | "type">, bytesBase64: string, operation: ExcelTransformRequest["operation"], mode: ExcelTransformRequest["mode"]) { const request = buildExcelTransformRequest(file, bytesBase64, operation, mode); mutate(request); return request; }

import { BlobByteSink } from "../../core/index.js";
import type { OutputDestination, ExportResult, ExportProgress } from "../../core/index.js";
import { writeXlsxTo } from "../../xlsx/index.js";
import { writeCsvTo } from "../../csv/index.js";
import { writePdfTo } from "../../pdf/index.js";
import { createCanopyExport } from "./source.js";
import type { CanopyCapture, CanopyFormat, CanopyWriteOptions } from "./types.js";
export { createCanopyExport } from "./source.js";
export { ExportCell } from "../../core/presentation.js";
export { PdfFont } from "../../pdf/font.js";
export type { CanopyCapture, CanopyExportCell, CanopyExportColumn, CanopyExportRow, CanopyExport, CanopyFormat, CanopyDiagnostic, CanopySourceOptions, CanopyWriteOptions } from "./types.js";

/** Write a captured CanopyX grid to a caller-owned destination; its partial bytes remain caller-owned on failure. */
export async function writeCanopyTo(capture: CanopyCapture, format: CanopyFormat, destination: OutputDestination, options: CanopyWriteOptions = {}): Promise<ExportResult> {
  const source = createCanopyExport(capture, format, options), { signal } = options;
  const progress = options.onProgress ? { onProgress: (event: ExportProgress) => {
    const result: unknown = options.onProgress!({ ...event, ...(source.rowCount === null ? {} : { totalRows: source.rowCount }) });
    if (result && typeof (result as { then?: unknown }).then === "function") { void Promise.resolve(result).catch(() => {}); throw new TypeError("onProgress must be synchronous."); }
  } } : {};
  const stream = { ...(signal ? { signal } : {}), ...progress };
  if (format === "xlsx") return writeXlsxTo(source.rows(), destination, { ...options.xlsx, ...stream, columns: source.columns,
    dateMode: options.xlsx?.dateMode ?? capture.request.timeZone,
    limits: { ...options.xlsx?.limits, ...options.limits } });
  if (format === "pdf") return writePdfTo(source.rows(), destination, { ...options.pdf, ...stream, columns: source.columns,
    limits: { ...options.pdf?.limits, ...options.limits } });
  return writeCsvTo(source.rows(), destination, { ...options.csv, ...stream, columns: source.columns, valueMode: capture.request.values,
    limits: { ...options.csv?.limits, ...options.limits } });
}

/** Export an immutable native capture as a Blob without loading or registering CanopyX. */
export async function exportCanopy(capture: CanopyCapture, format: CanopyFormat, options: CanopyWriteOptions = {}): Promise<Blob> {
  const sink = new BlobByteSink();
  try {
    await writeCanopyTo(capture, format, sink, options);
    return sink.toBlob(format === "xlsx" ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" : format === "pdf" ? "application/pdf" : "text/csv;charset=utf-8");
  } catch (error) { sink.discard(); throw error; }
}

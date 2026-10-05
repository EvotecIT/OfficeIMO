export { OfficeIMOError, NotSupportedError } from "./errors.js";
export type { ErrorCode } from "./errors.js";
export { BlobByteSink, ChunkedTextSink, writeBytes } from "./sinks.js";
export type { ByteSink, ByteSource, ByteWriter } from "./sinks.js";
export { checkAbort, withAbort, inputRows, pause } from "./iteration.js";
import { OfficeIMOError } from "./errors.js";

/** Plain values never interpret strings as formulas. */
export type CellValue = string | number | boolean | Date | null | undefined;
export type Row = readonly CellValue[] | Readonly<Record<string, CellValue>>;
export type Rows = Iterable<Row> | AsyncIterable<Row>;
export type Alignment = "left" | "center" | "right" | "fill" | "justify" | "distributed";

/** Shared ordered projection. Format modules interpret their own style options. */
export interface Column {
  readonly header: string;
  /** Dots are literal, never property traversal. */
  readonly key?: string;
  readonly width?: number;
  readonly type?: "string" | "number" | "boolean" | "date" | (string & {});
  readonly format?: string;
  readonly wrapText?: boolean;
  readonly alignment?: Alignment;
  /** XLSX StyleRegistry index. */
  readonly style?: number;
}

export interface ExportProgress {
  readonly phase: "rows" | "complete";
  readonly rows: number;
  readonly sheetName?: string;
  readonly bytes?: number;
}

export interface StreamOptions {
  readonly signal?: AbortSignal;
  /** Synchronous notification; throwing fails the operation. */
  readonly onProgress?: (progress: ExportProgress) => void;
}

/** Feature detection has no import-time DOM or compressor side effects. */
export function detectFeatures(): { blob: boolean; streams: boolean; deflateRaw: boolean; download: boolean } {
  let deflateRaw = false;
  if (typeof CompressionStream === "function") {
    try { new CompressionStream("deflate-raw"); deflateRaw = true; } catch { /* stored fallback */ }
  }
  return { blob: typeof Blob === "function", streams: typeof WritableStream === "function", deflateRaw,
    download: typeof document !== "undefined" && typeof URL.createObjectURL === "function" };
}

/** Browser download helper. Workers and Node should deliver the Blob themselves. */
export function saveBlob(blob: Blob, fileName: string): void {
  if (!(blob instanceof Blob) || typeof fileName !== "string" || !fileName) throw new TypeError("A Blob and file name are required.");
  if (typeof document === "undefined" || typeof URL.createObjectURL !== "function")
    throw new OfficeIMOError("PLATFORM_UNAVAILABLE", "saveBlob requires a browser document and object URLs.");
  const url = URL.createObjectURL(blob), link = document.createElement("a");
  try {
    link.href = url; link.download = fileName;
    document.body.append(link); link.click();
  } finally {
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  }
}

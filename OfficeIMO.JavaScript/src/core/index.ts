export { OfficeIMOError, NotSupportedError } from "./errors.js";
export type { ErrorCode } from "./errors.js";
export { BlobByteSink, ChunkedTextSink, writeBytes } from "./sinks.js";
export type { ByteSink, ByteSource, ByteWriter, OutputDestination } from "./sinks.js";
export type { ExportLimits } from "./limits.js";
export { ExportCell } from "./presentation.js";
export type { CellPresentation, ExportCellOptions, ExportValue } from "./presentation.js";
export type { ExportLink } from "./links.js";
export type { TableSpanCell, TableSpanRows } from "./table.js";
export { checkAbort, withAbort, inputRows, pause } from "./iteration.js";
import { OfficeIMOError } from "./errors.js";

/** Plain values never interpret strings as formulas. */
export type CellValue = string | number | boolean | Date | null | undefined;
export type Row = readonly import("./presentation.js").ExportValue[] | Readonly<Record<string, import("./presentation.js").ExportValue>>;
export type Rows = Iterable<Row> | AsyncIterable<Row>;
export type Alignment = "left" | "center" | "right" | "fill" | "justify" | "distributed";

/** Metadata shared by tabular writers. Format modules interpret their own presentation options. */
export interface ColumnSettings {
  readonly header: string;
  readonly width?: number;
  readonly type?: "string" | "number" | "boolean" | "date" | (string & {});
  readonly format?: string;
  readonly wrapText?: boolean;
  readonly alignment?: Alignment;
  /** Contiguous shared prefixes form merged heading rows above the leaf headers. */
  readonly groups?: readonly string[];
}

/** Zero-based position in the exported data, independent of titles and headings. */
export interface ColumnValueContext {
  readonly rowIndex: number;
  readonly columnIndex: number;
  readonly column: ColumnSettings & { readonly key?: string };
  readonly sheetName?: string;
  /** One-based Excel coordinate, when the destination is a worksheet. */
  readonly worksheetRow?: number;
}

/** Select a literal object key, or explicitly compute a value. Array rows remain positional unless every column has a getter. */
export type Column<T = never> = import("../internal/columns.js").ProjectedColumn<T, import("./presentation.js").ExportValue>;

/** Counts for one completed table export. Rows exclude generated headings and footers. */
export interface ExportResult {
  readonly rows: number;
  readonly columns: number;
  readonly bytes: number;
}

export interface ExportProgress {
  readonly phase: "rows" | "complete";
  readonly rows: number;
  readonly sheetName?: string;
  /** Completed rows in the named worksheet; rows remains the workbook-wide total. */
  readonly sheetRows?: number;
  readonly totalRows?: number;
  readonly bytes?: number;
}

export interface StreamOptions {
  readonly signal?: AbortSignal;
  readonly limits?: import("./limits.js").ExportLimits;
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

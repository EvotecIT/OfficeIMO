/** Tabular exports accept plain values and never interpret a string as a formula. */
export type CellValue = string | number | boolean | Date | null | undefined;
export type Row = readonly CellValue[] | Readonly<Record<string, CellValue>>;
export type Rows = Iterable<Row> | AsyncIterable<Row>;

export interface Column {
  /** Displayed header; also the default property key for object rows. */
  readonly header: string;
  /** Exact property key; dots are literal and never traverse an object. */
  readonly key?: string;
  /** XLSX width in characters, from 0 through 255. */
  readonly width?: number;
  /** Validates nonempty XLSX values; no string-to-number/date coercion. */
  readonly type?: "string" | "number" | "boolean" | "date";
  /** Excel number format code. Dates default to yyyy-mm-dd hh:mm:ss. */
  readonly format?: string;
  readonly wrapText?: boolean;
  readonly alignment?: "left" | "center" | "right" | "fill" | "justify" | "distributed";
}

export interface ExportProgress {
  readonly phase: "rows" | "complete";
  /** Consumed data rows, excluding the header. Per sheet during XLSX rows; total on completion. */
  readonly rows: number;
  readonly sheetName?: string;
  /** Final output size, available on completion. */
  readonly bytes?: number;
}

export interface StreamOptions {
  /** Cancels producer iteration, encoding and compression. Pass it to an I/O-bound producer too. */
  readonly signal?: AbortSignal;
  /** Synchronous notification; throwing fails the export. */
  readonly onProgress?: (progress: ExportProgress) => void;
}

/** Starts a download; browser/iframe download policy still applies. */
export declare function saveBlob(blob: Blob, fileName: string): void;

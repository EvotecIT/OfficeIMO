import type { CellValue, Column, Rows, StreamOptions } from "./common.js";
export { saveBlob } from "./common.js";
export type { CellValue, Column, Row, Rows, ExportProgress, StreamOptions } from "./common.js";

export interface WorkbookOptions extends StreamOptions {
  readonly creator?: string;
  readonly title?: string;
  readonly created?: Date;
  readonly modified?: Date;
  /** Date cell clock fields, default local. Workbook property dates are always UTC instants. */
  readonly dateMode?: "local" | "utc";
  /** Auto uses platform deflate-raw when supported; store always writes uncompressed ZIP entries. */
  readonly compression?: "auto" | "store";
}

export interface SheetOptions {
  readonly columns?: readonly Column[];
  readonly includeHeader?: boolean;
  readonly freezeHeader?: boolean;
  readonly autoFilter?: boolean;
  /** Default true. */
  readonly boldHeader?: boolean;
  /** RGB or ARGB hex, with an optional leading #. No fill by default. */
  readonly headerFill?: string;
}

export interface Sheet {
  /** Final sanitized, unique Excel sheet name. */
  readonly name: string;
  /** Consume once. Await before appending again or finalizing. A failed append invalidates the sheet. */
  addRows(rows: Rows): Promise<void>;
  /** Typed object records need no string index signature; every declared property must be a cell value. */
  addRows<T extends { readonly [K in keyof T]: CellValue }>(rows: Iterable<T> | AsyncIterable<T>): Promise<void>;
}

export interface Workbook {
  addSheet(name: string, options?: SheetOptions): Sheet;
  /** Finalize after all addRows calls complete. Repeated calls return the same output. */
  toBlob(): Promise<Blob>;
}

/** Creates a value-only XLSX workbook. Rows are encoded/compressed during addRows, not retained as strings. */
export declare function createWorkbook(options?: WorkbookOptions): Workbook;

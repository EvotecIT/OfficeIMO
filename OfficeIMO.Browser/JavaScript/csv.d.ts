import type { CellValue, Column, Rows, StreamOptions } from "./common.js";
export { saveBlob } from "./common.js";
export type { CellValue, Column, Row, Rows, ExportProgress, StreamOptions } from "./common.js";

export interface CsvOptions extends StreamOptions {
  readonly columns: readonly Column[];
  readonly delimiter?: "," | ";" | "\t";
  /** Default CRLF. */
  readonly lineEnding?: "\r\n" | "\n" | "\r";
  readonly bom?: boolean;
  readonly includeHeader?: boolean;
  /** Default true, using OfficeIMO.CSV's apostrophe-prefix rule. */
  readonly formulaInjectionProtection?: boolean;
}

/** Returns UTF-8 CSV, consuming rows once and yielding between bounded encoding batches. */
export declare function writeCsv(rows: Rows, options: CsvOptions): Promise<Blob>;
/** Accepts typed object records without requiring a string index signature. */
export declare function writeCsv<T extends { readonly [K in keyof T]: CellValue }>(rows: Iterable<T> | AsyncIterable<T>, options: CsvOptions): Promise<Blob>;

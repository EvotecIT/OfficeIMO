import { checkAbort, inputRows, withAbort } from "../core/iteration.js";
import { BlobByteSink, ChunkedTextSink } from "../core/sinks.js";
import type { ByteSink } from "../core/sinks.js";
import type { CellValue, Column, Rows, StreamOptions } from "../core/index.js";
import { copyColumns, rowValues } from "../internal/rows.js";
export { saveBlob } from "../core/index.js";
export type { CellValue, Column, Row, Rows, StreamOptions, ExportProgress } from "../core/index.js";

export interface CsvOptions extends StreamOptions {
  readonly columns: readonly Column[];
  readonly delimiter?: "," | ";" | "\t";
  readonly lineEnding?: "\r\n" | "\n" | "\r";
  readonly bom?: boolean;
  readonly includeHeader?: boolean;
  readonly formulaInjectionProtection?: boolean;
}
function csvField(value: unknown, delimiter: string, protect: boolean): string {
  let text: string;
  if (value == null) text = "";
  else if (value instanceof Date) text = Number.isFinite(value.getTime()) ? value.toISOString() : "";
  else if (typeof value === "boolean") text = value ? "True" : "False";
  else if (typeof value === "number") text = Number.isFinite(value) ? String(value) : "";
  else if (typeof value === "string") text = value;
  else throw new TypeError("CSV cells must be strings, numbers, booleans, Dates or null.");
  if (protect && typeof value === "string" && /^ *[=+\-@\t\r\n]/.test(text)) text = "'" + text;
  return text.includes(delimiter) || /["\r\n]/.test(text) ? '"' + text.replace(/"/g, '""') + '"' : text;
}

/** Stream UTF-8 to a caller-owned sink; a failing/cancelled destination owns partial-byte disposal. */
export function writeCsvTo(rows: Rows, sink: ByteSink, options: CsvOptions): Promise<void>;
export function writeCsvTo<T extends { readonly [K in keyof T]: CellValue }>(rows: Iterable<T> | AsyncIterable<T>, sink: ByteSink, options: CsvOptions): Promise<void>;
export async function writeCsvTo(rows: Iterable<unknown> | AsyncIterable<unknown>, sink: ByteSink, options: CsvOptions): Promise<void> {
  const columns = copyColumns(options.columns), delimiter = options.delimiter ?? ",", lineEnding = options.lineEnding ?? "\r\n";
  if (![",", ";", "\t"].includes(delimiter)) throw new RangeError("Delimiter must be comma, semicolon or tab.");
  if (!["\r\n", "\n", "\r"].includes(lineEnding)) throw new RangeError("Invalid line ending.");
  const signal = options.signal, protect = options.formulaInjectionProtection !== false, buffer = new ChunkedTextSink(sink, signal);
  checkAbort(signal);
  if (options.bom) await withAbort(Promise.resolve(sink.write(new Uint8Array([239, 187, 191]))), signal);
  let count = 0;
  async function record(values: readonly unknown[]): Promise<void> {
    for (let i = 0; i < columns.length; i++) await buffer.write((i ? delimiter : "") + csvField(values[i], delimiter, protect));
    if (buffer.append(lineEnding)) { await buffer.flush(); options.onProgress?.({ phase: "rows", rows: count }); }
  }
  if (options.includeHeader !== false && columns.length) await record(columns.map(c => c.header));
  for await (const row of inputRows(rows, signal)) { await record(rowValues(row, columns)); count++; }
  await buffer.close(); options.onProgress?.({ phase: "complete", rows: count }); checkAbort(signal);
}

export function writeCsv(rows: Rows, options: CsvOptions): Promise<Blob>;
export function writeCsv<T extends { readonly [K in keyof T]: CellValue }>(rows: Iterable<T> | AsyncIterable<T>, options: CsvOptions): Promise<Blob>;
export async function writeCsv(rows: Rows, options: CsvOptions): Promise<Blob> {
  const sink = new BlobByteSink();
  try {
    let completedRows = 0;
    await writeCsvTo(rows, sink, { ...options, onProgress: p => {
      if (p.phase === "complete") completedRows = p.rows;
      else options.onProgress?.(p);
    } });
    const blob = sink.toBlob("text/csv;charset=utf-8");
    options.onProgress?.({ phase: "complete", rows: completedRows, bytes: blob.size }); checkAbort(options.signal); return blob;
  } catch (error) { sink.discard(); throw error; }
}

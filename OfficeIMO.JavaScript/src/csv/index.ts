import { checkAbort, inputRows, withAbort } from "../core/iteration.js";
import { BlobByteSink, ChunkedTextSink } from "../core/sinks.js";
import { ExportBudget, boundedSink } from "../core/limits.js";
import { ExportCell } from "../core/presentation.js";
import type { ExportValue } from "../core/presentation.js";
import type { ByteSink } from "../core/sinks.js";
import type { CellValue, Column, Rows, StreamOptions } from "../core/index.js";
import { copyColumns, rowValues } from "../internal/rows.js";
export { saveBlob } from "../core/index.js";
export { ExportCell } from "../core/presentation.js";
export type { CellValue, Column, Row, Rows, StreamOptions, ExportProgress, ExportValue, CellPresentation } from "../core/index.js";

export interface CsvValueContext {
  /** One-based data row, excluding the header. */
  readonly row: number;
  readonly columnIndex: number;
  readonly column: Column;
  readonly values: readonly CellValue[];
}
export interface CsvColumn extends Column {
  /** Formatting happens before formula protection and RFC quoting. Return a scalar, never pre-escaped CSV. */
  readonly valueFormatter?: (value: CellValue, context: CsvValueContext) => CellValue;
}
export interface CsvOptions extends StreamOptions {
  readonly columns: readonly CsvColumn[];
  /** Raw typed values by default. Display mode uses ExportCell.text when supplied. */
  readonly valueMode?: "raw" | "display";
  readonly delimiter?: "," | ";" | "\t";
  readonly lineEnding?: "\r\n" | "\n" | "\r";
  readonly bom?: boolean;
  readonly includeHeader?: boolean;
  readonly formulaInjectionProtection?: boolean;
  readonly quote?: "minimal" | "all" | "strings";
  readonly nullValue?: string;
}
function csvField(value: unknown, delimiter: string, protect: boolean, quote: NonNullable<CsvOptions["quote"]>, nullValue?: string): string {
  if (value == null && nullValue !== undefined) value = nullValue;
  let text: string;
  if (value == null) text = "";
  else if (value instanceof Date) text = Number.isFinite(value.getTime()) ? value.toISOString() : "";
  else if (typeof value === "boolean") text = value ? "True" : "False";
  else if (typeof value === "number") text = Number.isFinite(value) ? String(value) : "";
  else if (typeof value === "string") text = value;
  else {
    if (typeof (value as { then?: unknown }).then === "function") void Promise.resolve(value).catch(() => {});
    throw new TypeError("CSV cells and formatter results must be synchronous strings, numbers, booleans, Dates or null.");
  }
  if (protect && typeof value === "string" && /^ *[=+\-@\t\r\n]/.test(text)) text = "'" + text;
  return quote === "all" || (quote === "strings" && typeof value === "string") || text.includes(delimiter) || /["\r\n]/.test(text)
    ? '"' + text.replace(/"/g, '""') + '"' : text;
}

/** Stream UTF-8 to a caller-owned sink; a failing/cancelled destination owns partial-byte disposal. */
export function writeCsvTo(rows: Rows, sink: ByteSink, options: CsvOptions): Promise<void>;
export function writeCsvTo<T extends { readonly [K in keyof T]: ExportValue }>(rows: Iterable<T> | AsyncIterable<T>, sink: ByteSink, options: CsvOptions): Promise<void>;
export async function writeCsvTo(rows: Iterable<unknown> | AsyncIterable<unknown>, sink: ByteSink, options: CsvOptions): Promise<void> {
  const columns = copyColumns(options.columns).map(c => Object.freeze(c)) as CsvColumn[], delimiter = options.delimiter ?? ",", lineEnding = options.lineEnding ?? "\r\n", quote = options.quote ?? "minimal";
  if (![",", ";", "\t"].includes(delimiter)) throw new RangeError("Delimiter must be comma, semicolon or tab.");
  if (!["\r\n", "\n", "\r"].includes(lineEnding)) throw new RangeError("Invalid line ending.");
  if (!["minimal", "all", "strings"].includes(quote)) throw new RangeError("Quoting must be minimal, all or strings.");
  if (options.valueMode !== undefined && !["raw", "display"].includes(options.valueMode)) throw new RangeError("valueMode must be raw or display.");
  if (options.nullValue !== undefined && typeof options.nullValue !== "string") throw new TypeError("nullValue must be a string.");
  for (const column of columns) if (column.valueFormatter !== undefined && typeof column.valueFormatter !== "function") throw new TypeError("CSV value formatters must be functions.");
  const budget = new ExportBudget(options.limits);
  sink = boundedSink(sink, budget);
  const hasFormatters = columns.some(column => column.valueFormatter);
  const signal = options.signal, protect = options.formulaInjectionProtection !== false, buffer = new ChunkedTextSink(sink, signal);
  checkAbort(signal);
  if (options.bom) await withAbort(Promise.resolve(sink.write(new Uint8Array([239, 187, 191]))), signal);
  let count = 0;
  async function record(values: readonly unknown[], header = false): Promise<void> {
    const resolved = (value: unknown) => value instanceof ExportCell ? options.valueMode === "display" && value.text !== undefined ? value.text : value.value : value;
    const snapshot = hasFormatters ? Object.freeze(columns.map((_, i) => resolved(values[i]))) as readonly CellValue[] : undefined;
    for (let i = 0; i < columns.length; i++) {
      const column = columns[i]!;
      const raw = resolved(values[i]);
      const value = !header && column.valueFormatter ? column.valueFormatter(raw as CellValue,
        { row: count + 1, columnIndex: i + 1, column, values: snapshot! }) : raw;
      budget.cell(value);
      if (buffer.append((i ? delimiter : "") + csvField(value, delimiter, protect, quote, options.nullValue))) await buffer.flush();
    }
    if (buffer.append(lineEnding)) { await buffer.flush(); options.onProgress?.({ phase: "rows", rows: count }); }
  }
  if (options.includeHeader !== false && columns.length) await record(columns.map(c => c.header), true);
  for await (const row of inputRows(rows, signal)) { budget.row(count + 1); await record(rowValues(row, columns)); count++; }
  await buffer.close(); options.onProgress?.({ phase: "complete", rows: count }); checkAbort(signal);
}

export function writeCsv(rows: Rows, options: CsvOptions): Promise<Blob>;
export function writeCsv<T extends { readonly [K in keyof T]: ExportValue }>(rows: Iterable<T> | AsyncIterable<T>, options: CsvOptions): Promise<Blob>;
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

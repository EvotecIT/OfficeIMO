import { beginTask, checkAbort, consumeRows, taskYieldDue, withAbort } from "../core/iteration.js";
import { BlobByteSink, ChunkedTextSink, withDestination } from "../core/sinks.js";
import { ExportBudget, boundedSink } from "../core/limits.js";
import { ExportCell } from "../core/presentation.js";
import { isDate } from "../core/dates.js";
import { notifyProgress } from "../core/progress.js";
import type { ByteSink, OutputDestination } from "../core/sinks.js";
import type { CellValue, Column, StreamOptions, ExportResult, TableSpanRows } from "../core/index.js";
import { copyColumns, createRowProjector } from "../internal/rows.js";
import { csvStructuredRows } from "./structured-rows.js";
import { captureBooleanText } from "../core/value-text.js";
import { captureCsvOptions } from "./options.js";
import type { BooleanText } from "../core/value-text.js";
export type { BooleanText } from "../core/value-text.js";
export { saveBlob } from "../core/index.js";
export { ExportCell } from "../core/presentation.js";
export type { CellValue, Column, ColumnValueContext, Row, Rows, StreamOptions, ExportProgress, ExportResult, OutputDestination, ExportValue, CellPresentation } from "../core/index.js";

export interface CsvValueContext {
  /** Zero-based data position, excluding the header. */
  readonly rowIndex: number;
  readonly columnIndex: number;
  readonly column: Column;
  readonly values: readonly CellValue[];
}
export type CsvColumn<T = never> = Column<T> & {
  readonly style?: never;
  /** Formatting happens before formula protection and RFC quoting. Return a scalar, never pre-escaped CSV. */
  readonly valueFormatter?: (value: CellValue, context: CsvValueContext) => CellValue;
};
export interface CsvOptions<T = never> extends StreamOptions {
  readonly columns: readonly CsvColumn<T>[];
  /** Raw typed values by default. Display mode uses ExportCell.text when supplied. */
  readonly valueMode?: "raw" | "display";
  readonly delimiter?: "," | ";" | "\t";
  readonly lineEnding?: "\r\n" | "\n" | "\r";
  readonly bom?: boolean;
  readonly includeHeader?: boolean;
  /** Replaces the ordinary leaf header; each physical matrix row is emitted. */
  readonly headerRows?: TableSpanRows;
  /** Appended report metadata, excluded from data-row progress and completion counts. */
  readonly footer?: { readonly rows: TableSpanRows };
  /** Required when an emitted header/footer spans cells. Repeat fills covered positions; anchors leaves them empty. */
  readonly spanMode?: "repeat" | "anchors";
  readonly formulaInjectionProtection?: boolean;
  readonly quote?: "minimal" | "all" | "strings";
  readonly nullValue?: string;
  /** Display policy only; defaults to the explicit compatibility strings True and False. */
  readonly booleanText?: BooleanText;
}
function csvField(value: unknown, delimiter: string, protect: boolean, quote: NonNullable<CsvOptions["quote"]>, booleanText: BooleanText, nullValue?: string): string {
  if (value == null && nullValue !== undefined) value = nullValue;
  let text: string;
  if (value == null) text = "";
  else if (isDate(value)) text = Number.isFinite(value.getTime()) ? value.toISOString() : "";
  else if (typeof value === "boolean") text = booleanText[value ? "true" : "false"];
  else if (typeof value === "number") text = Number.isFinite(value) ? String(value) : "";
  else if (typeof value === "string") text = value;
  else {
    if (typeof (value as { then?: unknown }).then === "function") void Promise.resolve(value).catch(() => {});
    throw new TypeError("CSV cells and formatter results must be synchronous strings, numbers, booleans, Dates or null.");
  }
  if (protect && (typeof value === "string" || typeof value === "boolean") && /^ *[=+\-@\t\r\n]/.test(text)) text = "'" + text;
  return quote === "all" || (quote === "strings" && typeof value === "string") || text.includes(delimiter) || /["\r\n]/.test(text)
    ? '"' + (text.includes('"') ? text.replace(/"/g, '""') : text) + '"' : text;
}

/** Stream UTF-8 to a caller-owned sink; a failing/cancelled destination owns partial-byte disposal. */
export function writeCsvTo<T extends object>(rows: Iterable<T> | AsyncIterable<T>, destination: OutputDestination, options: CsvOptions<NoInfer<T>>): Promise<ExportResult>;
export async function writeCsvTo(rows: Iterable<unknown> | AsyncIterable<unknown>, destination: OutputDestination, configuration: unknown): Promise<ExportResult> {
  const options = captureCsvOptions(configuration);
  return withDestination(destination, sink => write(rows, sink, options));
}
async function write(rows: Iterable<unknown> | AsyncIterable<unknown>, sink: ByteSink, options: CsvOptions): Promise<ExportResult> {
  beginTask();
  const columns = copyColumns(options.columns).map(c => Object.freeze(c)) as CsvColumn[], delimiter = options.delimiter ?? ",", lineEnding = options.lineEnding ?? "\r\n", quote = options.quote ?? "minimal";
  const booleanText = captureBooleanText(options.booleanText);
  if (![",", ";", "\t"].includes(delimiter)) throw new RangeError("Delimiter must be comma, semicolon or tab.");
  if (!["\r\n", "\n", "\r"].includes(lineEnding)) throw new RangeError("Invalid line ending.");
  if (!["minimal", "all", "strings"].includes(quote)) throw new RangeError("Quoting must be minimal, all or strings.");
  if (options.valueMode !== undefined && !["raw", "display"].includes(options.valueMode)) throw new RangeError("valueMode must be raw or display.");
  if (options.nullValue !== undefined && typeof options.nullValue !== "string") throw new TypeError("nullValue must be a string.");
  if (options.spanMode !== undefined && options.spanMode !== "repeat" && options.spanMode !== "anchors") throw new TypeError("CSV spanMode must be repeat or anchors.");
  const headings = options.includeHeader === false ? [] : options.headerRows === undefined ? columns.length ? [columns.map(c => c.header)] : [] : csvStructuredRows(options.headerRows, columns.length, options.spanMode);
  const footers = options.footer === undefined ? [] : csvStructuredRows(options.footer.rows, columns.length, options.spanMode);
  for (const column of columns) if (column.valueFormatter !== undefined && typeof column.valueFormatter !== "function") throw new TypeError("CSV value formatters must be functions.");
  const budget = new ExportBudget(options.limits);
  let bytes = 0;
  const accepted = sink;
  sink = boundedSink({ async write(chunk) { await accepted.write(chunk); bytes += chunk.byteLength; } }, budget);
  const project = createRowProjector(columns, undefined, options.signal);
  const hasFormatters = columns.some(column => column.valueFormatter);
  const signal = options.signal, protect = options.formulaInjectionProtection !== false, buffer = new ChunkedTextSink(sink, signal);
  checkAbort(signal);
  if (options.bom) await withAbort(Promise.resolve(sink.write(new Uint8Array([239, 187, 191]))), signal);
  let count = 0;
  const resolved = (value: unknown) => value !== null && typeof value === "object" && value instanceof ExportCell ? options.valueMode === "display" && value.text !== undefined ? value.text : value.value : value;
  function record(values: readonly unknown[], header = false, first = 0, snapshot?: readonly CellValue[]): void | Promise<void> {
    if (!header && hasFormatters && first === 0) snapshot = Object.freeze(columns.map((_, i) => resolved(values[i]))) as readonly CellValue[];
    let text = "";
    for (let i = first; i < columns.length; i++) {
      checkAbort(signal);
      const column = columns[i]!;
      const raw = resolved(values[i]);
      let value = !header && column.valueFormatter ? column.valueFormatter(raw as CellValue,
        { rowIndex: count, columnIndex: i, column, values: snapshot! }) : raw;
      if (value == null && options.nullValue !== undefined) value = options.nullValue;
      budget.cell(typeof value === "boolean" ? booleanText[value ? "true" : "false"] : value);
      checkAbort(signal);
      text += (i ? delimiter : "") + csvField(value, delimiter, protect, quote, booleanText, options.nullValue);
      if (text.length >= buffer.chunkSize || (i & 127) === 127 && taskYieldDue()) {
        if (buffer.append(text)) return buffer.flush().then(() => record(values, header, i + 1, snapshot));
        text = "";
      }
    }
    if (buffer.append(text + lineEnding)) return buffer.flush().then(() => { notifyProgress(options.onProgress, { phase: "rows", rows: count + (header ? 0 : 1) }); });
  }
  for (const heading of headings) await record(heading, true);
  await consumeRows(rows, signal, row => {
    budget.row(count + 1);
    const pending = record(project(row, count));
    if (pending) return pending.then(() => { count++; });
    count++;
  });
  for (const footer of footers) await record(footer, true);
  await buffer.close(); notifyProgress(options.onProgress, { phase: "complete", rows: count, bytes }); checkAbort(signal);
  return { rows: count, columns: columns.length, bytes };
}

export function writeCsv<T extends object>(rows: Iterable<T> | AsyncIterable<T>, options: CsvOptions<NoInfer<T>>): Promise<Blob>;
export async function writeCsv(rows: Iterable<unknown> | AsyncIterable<unknown>, configuration: unknown): Promise<Blob> {
  const options = captureCsvOptions(configuration);
  const sink = new BlobByteSink();
  try {
    let completedRows = 0;
    await write(rows, sink, { ...options, onProgress: p => {
      if (p.phase === "complete") completedRows = p.rows;
      else notifyProgress(options.onProgress, p);
    } });
    const blob = sink.toBlob("text/csv;charset=utf-8");
    notifyProgress(options.onProgress, { phase: "complete", rows: completedRows, bytes: blob.size }); checkAbort(options.signal); return blob;
  } catch (error) { sink.discard(); throw error; }
}

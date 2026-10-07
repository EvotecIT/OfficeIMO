import { withDestination } from "../core/sinks.js";
import { copyColumns } from "../internal/rows.js";
import type { Column, ExportResult, OutputDestination } from "../core/index.js";
import { Workbook } from "./workbook.js";
import type { WorkbookOptions, SheetOptions, XlsxRows } from "./types.js";

/** One table, using the same workbook engine as the advanced multi-worksheet API. */
export interface XlsxOptions<T = never> extends Omit<WorkbookOptions, "sink"> {
  readonly columns: readonly Column<T>[];
  readonly sheet?: Omit<SheetOptions, "columns" | "headerStyle"> & { readonly name?: string };
}

function prepare(options: XlsxOptions): XlsxOptions {
  const columns = copyColumns(options?.columns);
  if ((options.sheet as { headerStyle?: unknown } | undefined)?.headerStyle !== undefined)
    throw new TypeError("Workbook-local header styles require the advanced Workbook API; use boldHeader and headerFill.");
  return { ...options, columns };
}

function worksheet<T>(book: Workbook, options: XlsxOptions<T>) {
  const { name = "Data", ...sheet } = options.sheet ?? {};
  return book.addWorksheet(name, { boldHeader: true, autoFilter: sheet.includeHeader !== false,
    autoSize: { minWidth: 6, maxWidth: 54 }, ...sheet, columns: options.columns });
}

/** Write one worksheet to a Blob. Async sources are consumed once; call a source factory for each export. */
export function writeXlsx<T extends object>(rows: Iterable<T> | AsyncIterable<T>, options: XlsxOptions<NoInfer<T>>): Promise<Blob>;
export async function writeXlsx(rows: Iterable<unknown> | AsyncIterable<unknown>, configuration: unknown): Promise<Blob> {
  const options = prepare(configuration as XlsxOptions);
  const { columns: _columns, sheet: _sheet, ...settings } = options;
  const book = new Workbook(settings);
  try { await worksheet(book, options).addRows(rows as XlsxRows); return await book.toBlob(); }
  catch (error) { await book.discard(error); throw error; }
}

/** Await accepted bytes without closing/aborting the caller's destination. Its partial bytes remain caller-owned. */
export function writeXlsxTo<T extends object>(rows: Iterable<T> | AsyncIterable<T>, destination: OutputDestination, options: XlsxOptions<NoInfer<T>>): Promise<ExportResult>;
export async function writeXlsxTo(rows: Iterable<unknown> | AsyncIterable<unknown>, destination: OutputDestination, configuration: unknown): Promise<ExportResult> {
  const options = prepare(configuration as XlsxOptions);
  return withDestination(destination, async sink => {
    const { columns: _columns, sheet: _sheet, ...settings } = options;
    const book = new Workbook({ ...settings, sink });
    try {
      const sheet = worksheet(book, options);
      const columns = options.columns.length;
      await sheet.addRows(rows as XlsxRows);
      const result = await book.finish();
      return { rows: result.rows, columns, bytes: result.bytes };
    } catch (error) { await book.discard(error); throw error; }
  });
}

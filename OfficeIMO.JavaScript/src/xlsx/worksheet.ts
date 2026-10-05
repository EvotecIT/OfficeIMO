import { checkAbort, inputRows } from "../core/iteration.js";
import { ChunkedTextSink, BlobByteSink } from "../core/sinks.js";
import { NotSupportedError, OfficeIMOError } from "../core/errors.js";
import type { CellValue, Column } from "../core/index.js";
import { copyColumns, rowValues } from "../internal/rows.js";
import { EntryWriter } from "../zip/entry.js";
import type { PreparedEntry } from "../zip/entry.js";
import { xmlDeclaration } from "../xml/index.js";
import { Cell, cellText, columnName, inlineText, excelDate } from "./values.js";
import { spreadsheetNamespace, colorArgb } from "./styles.js";
import type { Workbook } from "./workbook.js";
import type { SheetOptions, XlsxRows } from "./types.js";

/** Worksheet rows are written once in order; the model retains compressed output rather than source data. */
export class Worksheet {
  private readonly columns: readonly Column[];
  private readonly declared: { column: Column; letter: string; style: number; dateStyle: number; headerStyle: number }[];
  private readonly options: SheetOptions;
  private entry: EntryWriter | undefined;
  private buffer: ChunkedTextSink | undefined;
  private readonly output = new BlobByteSink();
  private started = false;
  private busy = false;
  private failed = false;
  private error: unknown;
  private count = 0;
  private readonly headerRows: number;
  private constructor(private readonly book: Workbook, readonly name: string, options: SheetOptions) {
    this.columns = copyColumns(options.columns ?? []); this.options = { ...options };
    this.headerRows = this.columns.length && options.includeHeader !== false ? 1 : 0;
    this.declared = this.columns.map((column, i) => ({ column, letter: columnName(i + 1),
      style: book.styles.forColumn(column), dateStyle: book.styles.forColumn(column, false, undefined, true),
      headerStyle: book.styles.forColumn({ header: column.header, ...(column.wrapText === undefined ? {} : { wrapText: column.wrapText }),
        ...(column.alignment === undefined ? {} : { alignment: column.alignment }) }, options.boldHeader !== false, options.headerFill)
    }));
  }
  /** @internal */
  static create(book: Workbook, name: string, options: SheetOptions): Worksheet { return new Worksheet(book, name, options); }
  /** @internal Validate before allocating native compressor resources or registering the sheet name. */
  static validate(book: Workbook, options: SheetOptions): void {
    for (const feature of ["mergedCells", "hyperlinks", "conditionalFormats", "dataValidation"] as const)
      if (options[feature] !== undefined) throw new NotSupportedError(feature);
    const columns = copyColumns(options.columns ?? []);
    if (columns.length > 16384) throw new RangeError("Excel supports at most 16,384 columns.");
    if (options.includeHeader === false && (options.freezeHeader || options.autoFilter)) throw new TypeError("A frozen header or autofilter requires a header row.");
    if (options.headerFill !== undefined) colorArgb(options.headerFill);
    for (const c of columns) {
      if (c.width !== undefined && (!Number.isFinite(c.width) || c.width < 0 || c.width > 255)) throw new RangeError("Column width must be from 0 through 255 characters.");
      if (c.type !== undefined && !["string", "number", "boolean", "date"].includes(c.type) && !book.writerFor(c.type)) throw new TypeError("Invalid column type: " + c.type);
      if (c.alignment !== undefined && !["left", "center", "right", "fill", "justify", "distributed"].includes(c.alignment)) throw new TypeError("Invalid horizontal alignment.");
      cellText(c.header, book.settings.invalidCharacterPolicy);
      if (c.style !== undefined) book.styles.validateStyle(c.style);
      if (c.format !== undefined && (typeof c.format !== "string" || c.format.length > 255)) throw new TypeError("Column format must be a string of at most 255 characters.");
    }
  }
  private cell(value: unknown, i: number, row: number, header = false): string {
    const col = this.declared[i]!;
    const suppliedStyle = value instanceof Cell ? value.style : undefined;
    if (!header && col.column.type) {
      const writer = this.book.writerFor(col.column.type);
      if (writer) value = writer(value instanceof Cell ? value.value : value as CellValue,
        { column: Object.freeze({ ...col.column }), row, columnIndex: i + 1, sheetName: this.name });
    }
    const explicitStyle = value instanceof Cell ? value.style ?? suppliedStyle : suppliedStyle;
    if (value instanceof Cell) value = value.value;
    const type = value instanceof Date ? "date" : typeof value;
    if (!header && value != null && col.column.type && !this.book.writerFor(col.column.type) && col.column.type !== type)
      throw new TypeError("Cell " + col.letter + row + " does not match column type " + col.column.type + ".");
    const style = explicitStyle === undefined ? header ? col.headerStyle : type === "date" ? col.dateStyle : col.style : this.book.styles.validateStyle(explicitStyle);
    const prefix = '<c r="' + col.letter + row + '" s="' + style + '"';
    if (value == null || (type === "number" && !Number.isFinite(value))) return prefix + '/>';
    if (type === "date") value = excelDate(value as Date, this.book.settings.dateMode);
    if (value == null) return prefix + '/>';
    if (type === "string") return prefix + ' t="inlineStr"><is>' + inlineText(value, this.book.settings.invalidCharacterPolicy) + '</is></c>';
    if (type === "boolean") return prefix + ' t="b"><v>' + (value ? 1 : 0) + '</v></c>';
    if (type !== "number" && type !== "date") throw new TypeError("Excel cells must be strings, numbers, booleans, Dates or null.");
    return prefix + '><v>' + value + '</v></c>';
  }
  private async writeRow(values: readonly unknown[], number: number, header: boolean): Promise<void> {
    const buffer = this.buffer!;
    await buffer.write('<row r="' + number + '">');
    for (let i = 0; i < this.columns.length; i++) await buffer.write(this.cell(values[i], i, number, header));
    await buffer.write('</row>');
  }
  private async start(): Promise<void> {
    if (this.started) return;
    this.started = true;
    this.entry = new EntryWriter(this.book.settings.compression, this.output, this.book.settings.signal);
    const buffer = this.buffer = new ChunkedTextSink(this.entry, this.book.settings.signal), opts = this.options;
    await buffer.write(xmlDeclaration + '<worksheet xmlns="' + spreadsheetNamespace + '"><sheetViews><sheetView workbookViewId="0">' +
      (opts.freezeHeader && this.headerRows ? '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/>' : "") + '</sheetView></sheetViews>');
    if (this.columns.some(c => c.width !== undefined)) {
      await buffer.write('<cols>');
      for (let i = 0; i < this.columns.length; i++) {
        const column = this.columns[i]!;
        if (column.width !== undefined) await buffer.write('<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + column.width + '" customWidth="1"/>');
      }
      await buffer.write('</cols>');
    }
    await buffer.write('<sheetData>');
    if (this.headerRows) await this.writeRow(this.columns.map(c => c.header), 1, true);
  }
  addRows(rows: XlsxRows): Promise<void>;
  addRows<T extends { readonly [K in keyof T]: CellValue | Cell }>(rows: Iterable<T> | AsyncIterable<T>): Promise<void>;
  async addRows(rows: Iterable<unknown> | AsyncIterable<unknown>): Promise<void> {
    this.book.assertOpen();
    if (this.busy) throw new OfficeIMOError("INVALID_STATE", "Await the current addRows call before writing more rows to this sheet.");
    if (this.failed) throw this.error;
    this.busy = true;
    try {
      await this.start(); let checkpoint = performance.now();
      for await (const row of inputRows(rows, this.book.settings.signal)) {
        if (this.count + this.headerRows >= 1048576) throw new RangeError("Excel supports at most 1,048,576 rows including the header; split the sheet.");
        if (!this.columns.length) throw new RangeError("Declare columns before adding rows.");
        await this.writeRow(rowValues(row, this.columns), this.count + this.headerRows + 1, false); this.count++;
        if (performance.now() - checkpoint >= 50) { this.progress(); checkpoint = performance.now(); }
      }
      await this.buffer!.flush(); this.progress(); checkAbort(this.book.settings.signal);
    } catch (error) { this.failed = true; this.error = error; await this.discard(error); throw error; }
    finally { this.busy = false; }
  }
  private progress(): void { this.book.settings.onProgress?.({ phase: "rows", rows: this.count, sheetName: this.name }); }
  get rowCount(): number { return this.count; }
  /** @internal */
  get isBusy(): boolean { return this.busy; }
  /** @internal */
  async finish(): Promise<PreparedEntry> {
    if (this.failed) throw this.error;
    try {
      await this.start(); await this.buffer!.write('</sheetData>');
      if (this.options.autoFilter && this.headerRows) await this.buffer!.write('<autoFilter ref="A1:' + columnName(this.columns.length) + (this.count + 1) + '"/>');
      await this.buffer!.write('</worksheet>'); await this.buffer!.close();
      return { ...await this.entry!.close(), chunks: this.output.takeChunks() };
    } catch (error) { await this.discard(error); throw error; }
  }
  /** @internal */
  async discard(error: unknown): Promise<void> { await this.entry?.discard(error); this.output.discard(); }
}

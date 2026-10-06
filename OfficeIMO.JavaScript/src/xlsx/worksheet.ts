import { checkAbort, inputRows } from "../core/iteration.js";
import { ChunkedTextSink, BlobByteSink } from "../core/sinks.js";
import { NotSupportedError, OfficeIMOError } from "../core/errors.js";
import type { CellValue, Column } from "../core/index.js";
import { copyColumns, rowValues } from "../internal/rows.js";
import { EntryWriter } from "../zip/entry.js";
import type { PreparedEntry } from "../zip/entry.js";
import { xmlDeclaration } from "../xml/index.js";
import { officeRelationshipsNamespace } from "../opc/index.js";
import { Cell, cellText, columnName, inlineText, excelDate } from "./values.js";
import { spreadsheetNamespace, colorArgb } from "./styles.js";
import type { Workbook } from "./workbook.js";
import type { SheetOptions, XlsxRows } from "./types.js";
import type { RowStyleContext } from "./types.js";
import type { CellStyle } from "./styles.js";
import type { TableDefinition } from "./table.js";
import { copyHyperlink, copyImage, hyperlinksXml, cellPosition } from "./attachments.js";
import type { Hyperlink, WorksheetImage } from "./attachments.js";

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
  private readonly links: Hyperlink[] = [];
  private readonly pictures: WorksheetImage[] = [];
  private constructor(private readonly book: Workbook, readonly name: string, options: SheetOptions, private readonly table?: TableDefinition) {
    this.columns = copyColumns(options.columns ?? []);
    const alternate = options.alternatingRowStyle;
    this.options = { ...options, ...(alternate ? { alternatingRowStyle: { ...alternate,
      ...(typeof alternate.font === "object" ? { font: { ...alternate.font } } : {}),
      ...(typeof alternate.fill === "object" ? { fill: { ...alternate.fill } } : {}),
      ...(typeof alternate.border === "object" ? { border: Object.fromEntries(Object.entries(alternate.border).map(([side, edge]) => [side, { ...edge }])) } : {})
    } } : {}) };
    this.headerRows = this.columns.length && options.includeHeader !== false ? 1 : 0;
    for (const link of options.hyperlinks ?? []) this.links.push(copyHyperlink(link, book.settings.invalidCharacterPolicy));
    this.declared = this.columns.map((column, i) => ({ column, letter: columnName(i + 1),
      style: book.styles.forColumn(column), dateStyle: book.styles.forColumn(column, false, undefined, true),
      headerStyle: options.headerStyle ?? book.styles.forColumn({ header: column.header, ...(column.wrapText === undefined ? {} : { wrapText: column.wrapText }),
        ...(column.alignment === undefined ? {} : { alignment: column.alignment }) }, options.boldHeader !== false, options.headerFill)
    }));
  }
  /** @internal */
  static create(book: Workbook, name: string, options: SheetOptions, table?: TableDefinition): Worksheet { return new Worksheet(book, name, options, table); }
  /** @internal Validate before allocating native compressor resources or registering the sheet name. */
  static validate(book: Workbook, options: SheetOptions): void {
    for (const feature of ["mergedCells", "conditionalFormats", "dataValidation"] as const)
      if (options[feature] !== undefined) throw new NotSupportedError(feature);
    const columns = copyColumns(options.columns ?? []);
    if (options.hyperlinks !== undefined) {
      if (!Array.isArray(options.hyperlinks)) throw new TypeError("Hyperlinks must be an array.");
      const seen = new Set<string>();
      for (const requested of options.hyperlinks) {
        const link = copyHyperlink(requested, book.settings.invalidCharacterPolicy);
        if (seen.has(link.cell)) throw new TypeError("Duplicate hyperlink cell: " + link.cell);
        seen.add(link.cell);
      }
    }
    if (columns.length > 16384) throw new RangeError("Excel supports at most 16,384 columns.");
    if (options.includeHeader === false && (options.freezeHeader || options.autoFilter)) throw new TypeError("A frozen header or autofilter requires a header row.");
    if (options.headerFill !== undefined) colorArgb(options.headerFill);
    if (options.headerStyle !== undefined) book.styles.validateStyle(options.headerStyle);
    if (options.alternatingRowStyle !== undefined) book.styles.compose(0, options.alternatingRowStyle);
    for (const callback of [options.rowStyle, options.cellStyle]) if (callback !== undefined && typeof callback !== "function") throw new TypeError("Style callbacks must be functions.");
    for (const height of [options.rowHeight, options.headerHeight])
      if (height !== undefined && (!Number.isFinite(height) || height <= 0 || height > 409)) throw new RangeError("Row height must be positive and at most 409 points.");
    if (options.freezeColumns !== undefined && (!Number.isInteger(options.freezeColumns) || options.freezeColumns < 0 || options.freezeColumns > columns.length || options.freezeColumns >= 16384)) throw new RangeError("Frozen columns must be within the declared columns and leave a valid scrollable column.");
    if (options.table && options.includeHeader === false) throw new TypeError("Excel tables require a header row.");
    if (options.defaultColumnWidth !== undefined && (!Number.isFinite(options.defaultColumnWidth) || options.defaultColumnWidth < 0 || options.defaultColumnWidth > 255)) throw new RangeError("Default column width must be from 0 through 255 characters.");
    for (const c of columns) {
      if (c.width !== undefined && (!Number.isFinite(c.width) || c.width < 0 || c.width > 255)) throw new RangeError("Column width must be from 0 through 255 characters.");
      if (c.type !== undefined && !["string", "number", "boolean", "date"].includes(c.type) && !book.writerFor(c.type)) throw new TypeError("Invalid column type: " + c.type);
      if (c.alignment !== undefined && !["left", "center", "right", "fill", "justify", "distributed"].includes(c.alignment)) throw new TypeError("Invalid horizontal alignment.");
      cellText(c.header, book.settings.invalidCharacterPolicy);
      if (c.style !== undefined) book.styles.validateStyle(c.style);
      if (c.format !== undefined && (typeof c.format !== "string" || c.format.length > 255)) throw new TypeError("Column format must be a string of at most 255 characters.");
    }
  }
  private cell(value: unknown, i: number, row: number, header = false, rowStyle?: CellStyle, context?: RowStyleContext): string {
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
    let style = explicitStyle === undefined ? header ? col.headerStyle : type === "date" ? col.dateStyle : col.style : this.book.styles.validateStyle(explicitStyle);
    if (!header) {
      if (explicitStyle === undefined) {
        if ((row - this.headerRows) % 2 === 0 && this.options.alternatingRowStyle) style = this.book.styles.compose(style, this.options.alternatingRowStyle);
        if (rowStyle) style = this.book.styles.compose(style, rowStyle);
      }
      const patch = context && this.options.cellStyle?.({ ...context, value: value as CellValue, column: Object.freeze({ ...col.column }), columnIndex: i + 1 });
      if (patch) style = this.book.styles.compose(style, patch);
    }
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
    const height = header ? this.options.headerHeight : this.options.rowHeight;
    const context = !header ? Object.freeze({ row: number, sheetName: this.name,
      values: Object.freeze(this.columns.map((_, i) => values[i] instanceof Cell ? (values[i] as Cell).value : values[i])) as readonly CellValue[] }) : undefined;
    const rowStyle = context && this.options.rowStyle?.(context);
    await buffer.write('<row r="' + number + '"' + (height === undefined ? "" : ' ht="' + height + '" customHeight="1"') + '>');
    for (let i = 0; i < this.columns.length; i++) await buffer.write(this.cell(values[i], i, number, header, rowStyle, context));
    await buffer.write('</row>');
  }
  private async start(): Promise<void> {
    if (this.started) return;
    this.started = true;
    this.entry = new EntryWriter(this.book.settings.compression, this.output, this.book.settings.signal);
    const buffer = this.buffer = new ChunkedTextSink(this.entry, this.book.settings.signal), opts = this.options;
    const x = opts.freezeColumns ?? 0, y = opts.freezeHeader && this.headerRows ? 1 : 0;
    const pane = x && y ? "bottomRight" : x ? "topRight" : "bottomLeft", cell = columnName(x + 1) + (y + 1);
    await buffer.write(xmlDeclaration + '<worksheet xmlns="' + spreadsheetNamespace + '" xmlns:r="' + officeRelationshipsNamespace + '"><sheetViews><sheetView workbookViewId="0">' +
      (x || y ? '<pane' + (x ? ' xSplit="' + x + '"' : "") + (y ? ' ySplit="' + y + '"' : "") + ' topLeftCell="' + cell + '" activePane="' + pane + '" state="frozen"/><selection pane="' + pane + '" activeCell="' + cell + '" sqref="' + cell + '"/>' : "") + '</sheetView></sheetViews>');
    const defaultWidth = opts.defaultColumnWidth ?? (this.table ? 20 : undefined);
    if (this.columns.some(c => c.width !== undefined || defaultWidth !== undefined)) {
      await buffer.write('<cols>');
      for (let i = 0; i < this.columns.length; i++) {
        const column = this.columns[i]!;
        const width = column.width ?? defaultWidth;
        if (width !== undefined) await buffer.write('<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + width + '" customWidth="1"/>');
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
  /** Register a PNG such as a chart; bytes are copied so the caller can reuse its buffer. */
  addImage(image: WorksheetImage): void {
    this.book.assertOpen();
    if (this.failed) throw this.error;
    this.pictures.push(copyImage(image, this.book.settings.invalidCharacterPolicy));
  }
  /** Attach an external report link to a cell without changing its literal value. */
  addHyperlink(link: Hyperlink): void {
    this.book.assertOpen();
    if (this.failed) throw this.error;
    const copied = copyHyperlink(link, this.book.settings.invalidCharacterPolicy);
    if (this.links.some(existing => existing.cell === copied.cell)) throw new TypeError("Duplicate hyperlink cell: " + copied.cell);
    this.links.push(copied);
  }
  /** @internal */
  get hyperlinks(): readonly Hyperlink[] { return this.links; }
  /** @internal */
  get images(): readonly WorksheetImage[] { return this.pictures; }
  /** @internal An empty export has no data table; it still emits the declared headers. */
  get tableDefinition(): TableDefinition | undefined { return this.count ? this.table : undefined; }
  /** @internal */
  get isBusy(): boolean { return this.busy; }
  /** @internal */
  async finish(): Promise<PreparedEntry> {
    if (this.failed) throw this.error;
    try {
      for (const link of this.links) {
        const position = cellPosition(link.cell);
        if (position.column > this.columns.length || position.row > this.count + this.headerRows) throw new RangeError("Hyperlinks must address cells within the exported rows and columns.");
      }
      await this.start(); await this.buffer!.write('</sheetData>');
      if (this.options.autoFilter && this.headerRows && !this.tableDefinition) await this.buffer!.write('<autoFilter ref="A1:' + columnName(this.columns.length) + (this.count + 1) + '"/>');
      if (this.links.length) await this.buffer!.write(hyperlinksXml(this.links, this.book.settings.invalidCharacterPolicy));
      if (this.pictures.length) await this.buffer!.write('<drawing r:id="drawing"/>');
      if (this.tableDefinition) await this.buffer!.write('<tableParts count="1"><tablePart r:id="table"/></tableParts>');
      await this.buffer!.write('</worksheet>'); await this.buffer!.close();
      return { ...await this.entry!.close(), chunks: this.output.takeChunks() };
    } catch (error) { await this.discard(error); throw error; }
  }
  /** @internal */
  async discard(error: unknown): Promise<void> { await this.entry?.discard(error); this.output.discard(); }
}

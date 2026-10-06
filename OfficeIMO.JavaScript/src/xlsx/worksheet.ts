import { checkAbort, inputRows } from "../core/iteration.js";
import { ChunkedTextSink, BlobByteSink } from "../core/sinks.js";
import { NotSupportedError, OfficeIMOError } from "../core/errors.js";
import type { CellValue, Column } from "../core/index.js";
import { copyColumns, rowValues } from "../internal/rows.js";
import { EntryWriter } from "../zip/entry.js";
import type { ZipEntry } from "../zip/index.js";
import type { PreparedEntry } from "../zip/entry.js";
import { xmlDeclaration } from "../xml/index.js";
import { officeRelationshipsNamespace } from "../opc/index.js";
import { Cell, cellText, columnName, inlineText, excelDate, copyValue } from "./values.js";
import { spreadsheetNamespace, colorArgb, validateStylePatch, copyStylePatch } from "./styles.js";
import type { Workbook } from "./workbook.js";
import type { SheetOptions, XlsxRows } from "./types.js";
import type { RowStyleContext } from "./types.js";
import type { CellStyle } from "./styles.js";
import type { TableDefinition } from "./table.js";
import { copyHyperlink, copyImage, hyperlinksXml, cellPosition } from "./attachments.js";
import type { Hyperlink, WorksheetImage } from "./attachments.js";
import { ExportCell, assertScalar } from "../core/presentation.js";
import type { ExportValue } from "../core/presentation.js";
import { ReportLayout, ComputedTotal, printXml } from "./layout.js";
import { cleanXml } from "../xml/index.js";

/** Worksheet rows are written once in order; the model retains compressed output rather than source data. */
export class Worksheet {
  private readonly columns: readonly Column[];
  private readonly declared: { column: Column; letter: string; style: number; dateStyle: number; headerStyle: number }[];
  private readonly options: SheetOptions;
  private entry: EntryWriter | ZipEntry | undefined;
  private completion: Promise<void> | undefined;
  private prepared: PreparedEntry | undefined;
  private buffer: ChunkedTextSink | undefined;
  private readonly output = new BlobByteSink();
  private started = false;
  private busy = false;
  private failed = false;
  private error: unknown;
  private count = 0;
  private readonly headerRows: number;
  private readonly layout: ReportLayout;
  private pending: string[][] = [];
  private pendingCharacters = 0;
  private reservedLayout = false;
  private readonly internalLinks: { cell: string; location: string }[] = [];
  private readonly links: Hyperlink[] = [];
  private readonly pictures: WorksheetImage[] = [];
  private constructor(private readonly book: Workbook, readonly name: string, options: SheetOptions, private readonly table?: TableDefinition, private readonly preserved = false) {
    this.columns = copyColumns(options.columns ?? []).map(c => Object.freeze(c));
    const alternate = options.alternatingRowStyle;
    this.options = { ...options, ...(options.autoSize ? { autoSize: { ...options.autoSize } } : {}),
      ...(options.print ? { print: { ...options.print, ...(options.print.margins ? { margins: { ...options.print.margins } } : {}) } } : {}),
      ...(options.footer ? { footer: { ...options.footer, ...(options.footer.values ? { values: options.footer.values.map(copyValue) as NonNullable<SheetOptions["footer"]>["values"] & {} } : {}), ...(options.footer.totals ? { totals: { ...options.footer.totals } } : {}), ...(options.footer.style ? { style: copyStylePatch(options.footer.style) } : {}) } } : {}),
      ...(alternate ? { alternatingRowStyle: copyStylePatch(alternate) } : {}) };
    this.layout = new ReportLayout(this.columns, this.options, book.settings.invalidCharacterPolicy);
    this.headerRows = this.layout.headerRows;
    for (const link of options.hyperlinks ?? []) { book.retainLink(); this.links.push(copyHyperlink(link, book.settings.invalidCharacterPolicy)); }
    this.declared = this.columns.map((column, i) => ({ column, letter: columnName(i + 1),
      style: book.styles.forColumn(column), dateStyle: book.styles.forColumn(column, false, undefined, true),
      headerStyle: this.headerRows ? options.headerStyle ?? book.styles.forColumn({ header: column.header, ...(column.wrapText === undefined ? {} : { wrapText: column.wrapText }),
        ...(column.alignment === undefined ? {} : { alignment: column.alignment }) }, options.boldHeader !== false, options.headerFill) : 0
    }));
  }
  /** @internal */
  static create(book: Workbook, name: string, options: SheetOptions, table?: TableDefinition, preserved = false): Worksheet { return new Worksheet(book, name, options, table, preserved); }
  /** @internal Validate before allocating native compressor resources or registering the sheet name. */
  static validate(book: Workbook, options: SheetOptions): void {
    for (const feature of ["mergedCells", "conditionalFormats", "dataValidation"] as const)
      if (options[feature] !== undefined) throw new NotSupportedError(feature);
    const columns = copyColumns(options.columns ?? []);
    new ReportLayout(columns, options, book.settings.invalidCharacterPolicy);
    if (options.footer?.style) book.styles.compose(0, options.footer.style);
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
  private cell(value: unknown, i: number, row: number, header = false, rowStyle?: CellStyle, context?: RowStyleContext, rowStyles?: Map<number, number>, footer = false): string {
    const col = this.declared[i]!;
    const total = value instanceof ComputedTotal ? value : undefined;
    if (total) value = total.value;
    let presentation = value instanceof ExportCell ? value.presentation : undefined;
    if (value instanceof ExportCell) value = value.value;
    const suppliedStyle = value instanceof Cell ? value.style : undefined;
    if (!header && !footer && col.column.type) {
      const writer = this.book.writerFor(col.column.type);
      if (writer) value = writer(value instanceof Cell ? value.value : value as CellValue,
        { column: col.column, row, columnIndex: i + 1, sheetName: this.name });
    }
    if (value instanceof ExportCell) { presentation = value.presentation; value = value.value; }
    const explicitStyle = value instanceof Cell ? value.style ?? suppliedStyle : suppliedStyle;
    if (value instanceof Cell) value = value.value;
    assertScalar(value);
    const type = value instanceof Date ? "date" : typeof value;
    if (!header && !footer && value != null && col.column.type && !this.book.writerFor(col.column.type) && col.column.type !== type)
      throw new TypeError("Cell " + col.letter + row + " does not match column type " + col.column.type + ".");
    let style = explicitStyle === undefined ? header ? col.headerStyle : total?.operation === "count" ? 0 : type === "date" ? col.dateStyle : col.style : this.book.styles.validateStyle(explicitStyle);
    if (!header && !footer) {
      if (explicitStyle === undefined) {
        const base = style, cached = rowStyles?.get(base);
        if (cached !== undefined) style = cached;
        else {
          if ((row - this.headerRows) % 2 === 0 && this.options.alternatingRowStyle) style = this.book.styles.compose(style, this.options.alternatingRowStyle);
          if (rowStyle !== undefined) style = this.book.styles.compose(style, rowStyle);
          rowStyles?.set(base, style);
        }
      }
      const patch = context && this.options.cellStyle?.({ ...context, value: value as CellValue, column: col.column, columnIndex: i + 1 });
      if (patch !== undefined) style = this.book.styles.compose(style, patch);
    }
    if (footer && this.options.footer?.style && explicitStyle === undefined) style = this.book.styles.compose(style, this.options.footer.style);
    if (presentation && explicitStyle === undefined) style = this.book.styles.compose(style, {
      ...(presentation.background === undefined ? {} : { fill: { color: presentation.background } }),
      ...(presentation.color === undefined && presentation.bold === undefined && presentation.italic === undefined ? {} : { font: {
        ...(presentation.color === undefined ? {} : { color: presentation.color }), ...(presentation.bold === undefined ? {} : { bold: presentation.bold }), ...(presentation.italic === undefined ? {} : { italic: presentation.italic }) } }),
      ...(presentation.wrapText === undefined ? {} : { wrapText: presentation.wrapText }), ...(presentation.alignment === undefined ? {} : { alignment: presentation.alignment }),
      ...(presentation.numberFormat === undefined ? {} : { numberFormat: presentation.numberFormat })
    });
    const originalCharacters = typeof value === "string" ? value.length : 0;
    if (type === "string" && cleanXml(value, this.book.settings.invalidCharacterPolicy).length > 32767) {
      const cell = col.letter + row;
      if (this.links.some(link => link.cell === cell)) throw new TypeError("A text-preservation cell cannot also have an external hyperlink.");
      const preserved = this.book.preserveText(this.name, cell, value as string);
      this.book.retainLink(); this.internalLinks.push({ cell, location: preserved.location }); value = preserved.preview;
    }
    if (type === "date") value = excelDate(value as Date, this.book.settings.dateMode);
    if (!header && !footer) this.layout.accept(i, value as CellValue);
    const prefix = '<c r="' + col.letter + row + '" s="' + style + '"';
    this.book.budget.cell(value, this.preserved || header || footer ? originalCharacters : undefined);
    if (total) return prefix + (value === null ? ' t="str"' : "") + '><f>' + total.formula.replace(/"/g, "&quot;") + '</f><v>' + (value ?? "") + '</v></c>';
    if (value == null || (type === "number" && !Number.isFinite(value))) return prefix + '/>';
    if (value == null) return prefix + '/>';
    if (type === "string") return prefix + ' t="inlineStr"><is>' + inlineText(value, this.book.settings.invalidCharacterPolicy) + '</is></c>';
    if (type === "boolean") return prefix + ' t="b"><v>' + (value ? 1 : 0) + '</v></c>';
    if (type !== "number" && type !== "date") throw new TypeError("Excel cells must be strings, numbers, booleans, Dates or null.");
    return prefix + '><v>' + value + '</v></c>';
  }
  private *rowXml(values: readonly unknown[], number: number, header: boolean, footer = false): Generator<string> {
    const height = header ? this.options.headerHeight : this.options.rowHeight;
    const context = !header && !footer && (this.options.rowStyle || this.options.cellStyle) ? Object.freeze({ row: number, sheetName: this.name,
      values: Object.freeze(this.columns.map((_, i) => values[i] instanceof Cell || values[i] instanceof ExportCell ? (values[i] as Cell | ExportCell).value : values[i])) as readonly CellValue[] }) : undefined;
    const rowStyle = context && this.options.rowStyle?.(context);
    if (rowStyle !== undefined) validateStylePatch(rowStyle);
    const rowStyles = !header && (rowStyle || this.options.alternatingRowStyle) ? new Map<number, number>() : undefined;
    yield '<row r="' + number + '"' + (height === undefined ? "" : ' ht="' + height + '" customHeight="1"') + '>';
    for (let i = 0; i < this.columns.length; i++) yield this.cell(values[i], i, number, header, rowStyle, context, rowStyles, footer);
    yield '</row>';
  }
  private async writeRow(values: readonly unknown[], number: number, header: boolean, footer = false): Promise<void> {
    for (const chunk of this.rowXml(values, number, header, footer)) if (this.buffer!.append(chunk)) await this.buffer!.flush();
  }
  private reserveLayout(): void {
    if (this.reservedLayout) return;
    if (!this.preserved) {
      let characters = 0;
      if (this.headerRows) {
        for (const heading of this.layout.headings) for (const text of heading) characters += text.length;
        for (const column of this.columns) characters += column.header.length;
      }
      if (this.options.footer) for (let i = 0; i < this.columns.length; i++) {
        const column = this.columns[i]!, totals = this.options.footer.totals;
        if (totals && Object.prototype.hasOwnProperty.call(totals, column.key ?? column.header)) continue;
        const value = this.options.footer.values?.[i], raw = value instanceof ExportCell || value instanceof Cell ? value.value : value;
        if (typeof raw === "string") characters += raw.length;
      }
      this.book.budget.reserve(this.columns.length * (this.headerRows + (this.options.footer ? 1 : 0)), characters);
    }
    this.reservedLayout = true;
  }
  private async start(): Promise<void> {
    if (this.started) return;
    this.reserveLayout();
    this.started = true;
    this.entry = this.book.settings.sink ? await this.book.openSheet(this) : new EntryWriter(this.book.settings.compression, { write: bytes => { this.book.retainBufferedBytes(bytes.length); this.output.write(bytes); } }, this.book.settings.signal);
    const buffer = this.buffer = new ChunkedTextSink(this.entry, this.book.settings.signal), opts = this.options;
    const x = opts.freezeColumns ?? 0, y = opts.freezeHeader ? this.headerRows : 0;
    const pane = x && y ? "bottomRight" : x ? "topRight" : "bottomLeft", cell = columnName(x + 1) + (y + 1);
    await buffer.write(xmlDeclaration + '<worksheet xmlns="' + spreadsheetNamespace + '" xmlns:r="' + officeRelationshipsNamespace + '">' + (opts.print ? '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' : "") + '<sheetViews><sheetView workbookViewId="0">' +
      (x || y ? '<pane' + (x ? ' xSplit="' + x + '"' : "") + (y ? ' ySplit="' + y + '"' : "") + ' topLeftCell="' + cell + '" activePane="' + pane + '" state="frozen"/><selection pane="' + pane + '" activeCell="' + cell + '" sqref="' + cell + '"/>' : "") + '</sheetView></sheetViews>');
    const defaultWidth = opts.defaultColumnWidth ?? (this.table ? 20 : undefined);
    if (this.columns.some((_, i) => this.layout.widths[i] !== undefined || defaultWidth !== undefined)) {
      await buffer.write('<cols>');
      for (let i = 0; i < this.columns.length; i++) {
        const width = this.layout.widths[i] ?? defaultWidth;
        if (width !== undefined) await buffer.write('<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + width + '" customWidth="1"/>');
      }
      await buffer.write('</cols>');
    }
    await buffer.write('<sheetData>');
    if (this.headerRows) {
      for (let i = 0; i < this.layout.headings.length; i++) await this.writeRow(this.layout.headings[i]!, i + 1, true);
      await this.writeRow(this.columns.map(c => c.header), this.headerRows, true);
    }
    for (const row of this.pending) for (const chunk of row) if (buffer.append(chunk)) await buffer.flush();
    this.pending = []; this.pendingCharacters = 0;
  }
  addRows(rows: XlsxRows): Promise<void>;
  addRows<T extends { readonly [K in keyof T]: ExportValue | Cell }>(rows: Iterable<T> | AsyncIterable<T>): Promise<void>;
  async addRows(rows: Iterable<unknown> | AsyncIterable<unknown>): Promise<void> {
    this.book.assertOpen();
    if (this.completion) throw new OfficeIMOError("INVALID_STATE", "Worksheet is closed.");
    if (this.busy) throw new OfficeIMOError("INVALID_STATE", "Await the current addRows call before writing more rows to this sheet.");
    if (this.failed) throw this.error;
    this.busy = true;
    try {
      this.reserveLayout();
      if (!this.layout.sampleRows) await this.start();
      let checkpoint = performance.now();
      for await (const row of inputRows(rows, this.book.settings.signal)) {
        if (this.count + this.headerRows + (this.options.footer ? 1 : 0) >= 1048576) throw new RangeError("Excel supports at most 1,048,576 rows including headers and footers; split the sheet.");
        this.book.budget.row(this.count + 1);
        if (!this.columns.length) throw new RangeError("Declare columns before adding rows.");
        const values = rowValues(row, this.columns);
        if (!this.started) {
          if ((this.pending.length + 1) * this.columns.length > (this.book.settings.limits?.maxBufferedCells ?? 100000)) throw new OfficeIMOError("RESOURCE_LIMIT", "Width sampling cell limit exceeded; reduce sampleRows or raise its bounded limit.");
          const encoded: string[] = [];
          for (const chunk of this.rowXml(values, this.count + this.headerRows + 1, false)) {
            if (this.pendingCharacters + chunk.length > (this.book.settings.limits?.maxBufferedCharacters ?? 1000000)) throw new OfficeIMOError("RESOURCE_LIMIT", "Width sampling text limit exceeded; reduce sampleRows or raise its bounded limit.");
            this.pendingCharacters += chunk.length; encoded.push(chunk);
          }
          this.layout.sample(values);
          this.pending.push(encoded); this.count++;
          if (this.pending.length >= this.layout.sampleRows) await this.start();
        } else { await this.writeRow(values, this.count + this.headerRows + 1, false); this.count++; }
        if (performance.now() - checkpoint >= 50) { this.progress(); checkpoint = performance.now(); }
      }
      if (this.buffer) await this.buffer.flush(); this.progress(); checkAbort(this.book.settings.signal);
    } catch (error) { this.failed = true; this.error = error; await this.book.discard(error); throw error; }
    finally { this.busy = false; }
  }
  private progress(): void { this.book.settings.onProgress?.({ phase: "rows", rows: this.count, sheetName: this.name }); }
  get rowCount(): number { return this.count; }
  /** Register a PNG such as a chart; bytes are copied so the caller can reuse its buffer. */
  addImage(image: WorksheetImage): void {
    this.book.assertOpen();
    if (this.completion) throw new OfficeIMOError("INVALID_STATE", "Worksheet is closed.");
    if (this.failed) throw this.error;
    const copied = copyImage(image, this.book.settings.invalidCharacterPolicy);
    this.book.retainImage(copied.data.length); this.pictures.push(copied);
  }
  /** Attach an external report link to a cell without changing its literal value. */
  addHyperlink(link: Hyperlink): void {
    this.book.assertOpen();
    if (this.completion) throw new OfficeIMOError("INVALID_STATE", "Worksheet is closed.");
    if (this.failed) throw this.error;
    const copied = copyHyperlink(link, this.book.settings.invalidCharacterPolicy);
    if (this.links.some(existing => existing.cell === copied.cell) || this.internalLinks.some(existing => existing.cell === copied.cell)) throw new TypeError("Duplicate hyperlink cell: " + copied.cell);
    this.book.retainLink();
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
  get headerRowCount(): number { return this.headerRows; }
  /** @internal */
  get totalRows(): number { return Math.max(1, this.headerRows + this.count + (this.options.footer ? 1 : 0)); }
  /** @internal */
  get lastColumn(): string { return columnName(Math.max(1, this.columns.length)); }
  /** @internal */
  get printSettings(): SheetOptions["print"] { return this.options.print; }
  /** @internal */
  get footerSettings(): SheetOptions["footer"] { return this.options.footer; }
  /** Complete this worksheet. A streamed workbook can then start its next worksheet. */
  async close(): Promise<void> {
    this.book.assertOpen();
    if (this.busy) throw new OfficeIMOError("INVALID_STATE", "Await addRows before closing the worksheet.");
    try { await this.finish(); } catch (error) { await this.book.discard(error); throw error; }
  }
  /** @internal */
  finish(preservedRows?: XlsxRows): Promise<void> {
    if (this.completion) return this.completion;
    if (this.failed) throw this.error;
    return this.completion = (async () => { try {
      for (const link of this.links) {
        const position = cellPosition(link.cell);
        if (position.column > this.columns.length || position.row > this.totalRows) throw new RangeError("Hyperlinks must address cells within the exported rows and columns.");
      }
      await this.start();
      if (preservedRows) for await (const row of inputRows(preservedRows, this.book.settings.signal)) {
        this.book.budget.row(this.count + 1); await this.writeRow(rowValues(row, this.columns), this.headerRows + this.count + 1, false); this.count++;
      }
      if (this.options.footer) await this.writeRow(this.layout.footer(this.count), this.headerRows + this.count + 1, false, true);
      await this.buffer!.write('</sheetData>');
      if (this.options.autoFilter && this.headerRows && !this.tableDefinition) await this.buffer!.write('<autoFilter ref="A' + this.headerRows + ':' + columnName(this.columns.length) + (this.count + this.headerRows) + '"/>');
      if (this.layout.merges.length) await this.buffer!.write('<mergeCells count="' + this.layout.merges.length + '">' + this.layout.merges.map(ref => '<mergeCell ref="' + ref + '"/>').join("") + '</mergeCells>');
      if (this.links.length || this.internalLinks.length) await this.buffer!.write(hyperlinksXml(this.links, this.book.settings.invalidCharacterPolicy, this.internalLinks));
      await this.buffer!.write(printXml(this.options.print, this.book.settings.invalidCharacterPolicy));
      if (this.pictures.length) await this.buffer!.write('<drawing r:id="drawing"/>');
      if (this.tableDefinition) await this.buffer!.write('<tableParts count="1"><tablePart r:id="table"/></tableParts>');
      await this.buffer!.write('</worksheet>'); await this.buffer!.close();
      const info = await this.entry!.close();
      if (!this.book.settings.sink && !this.failed) this.prepared = { ...info, chunks: this.output.takeChunks() };
    } catch (error) { await this.discard(error); throw error; } })();
  }
  /** @internal */
  takePrepared(): PreparedEntry | undefined { const prepared = this.prepared; this.prepared = undefined; return prepared; }
  /** @internal */
  async discard(error: unknown): Promise<void> { this.failed = true; this.error = error; this.pending = []; this.pendingCharacters = 0; this.prepared = undefined; this.buffer = undefined; await this.entry?.discard(error); this.output.discard(); }
}

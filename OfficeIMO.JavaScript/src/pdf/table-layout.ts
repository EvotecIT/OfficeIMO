import { ExportCell } from "../core/presentation.js";
import { copyExportLink } from "../core/links.js";
import type { ExportLink } from "../core/links.js";
import type { ExportValue, CellPresentation, TableSpanRows } from "../core/index.js";
import { tableSpans } from "../internal/table-spans.js";
import { createTotals } from "../internal/totals.js";
import type { NumericAggregate } from "../internal/totals.js";
import type { PdfFontResources, PdfFontResource } from "./font-resources.js";
import type { PdfSettings } from "./settings.js";
import { presentation, positive } from "./settings.js";
import { displayText } from "./text.js";
import { capturePdfText, wrapPdfText } from "./inline-text.js";
import { planTableBands, projectBandRows, alignBandBlocks } from "./table-band-plan.js";
import type { PdfTableBand } from "./table-band-plan.js";
import type { PdfCapturedText } from "./inline-text.js";
import type { PdfInlineLine } from "./inline-text.js";
import type { PdfText } from "./inline-types.js";

export interface PdfLayoutCell {
  readonly first: number;
  readonly span: number;
  readonly width: number;
  readonly font: PdfFontResource;
  readonly style: Readonly<CellPresentation>;
  readonly lines: readonly PdfInlineLine[];
  readonly lineHeight: number;
  readonly rowSpan?: number;
  readonly height?: number;
  readonly link?: Readonly<ExportLink>;
  /** Captured once for band span projection; ordinary layouts do not retain it. */
  readonly captured?: PdfCapturedText;
}
export interface PdfLayoutRow { readonly cells: readonly PdfLayoutCell[]; readonly lines: number; readonly lineHeight: number; readonly height: number; readonly structured?: boolean; }
interface HeadingCell { readonly first: number; readonly span: number; readonly text: string; }
export class PdfTableLayout {
  readonly bands: readonly PdfTableBand[] | undefined;
  readonly widths: readonly number[];
  readonly lefts: readonly number[];
  readonly lineHeight: number;
  readonly headings: readonly (readonly HeadingCell[])[];
  readonly totals: readonly NumericAggregate[];
  private readonly repeatedColumns: ReadonlySet<number>;
  private readonly structuredBandRows = new WeakMap<readonly PdfLayoutRow[], Promise<readonly (readonly PdfLayoutRow[])[]>>();
  constructor(readonly settings: PdfSettings, readonly fonts: PdfFontResources) {
    const { options, padding, fontSize, page, margins } = settings, columns = options.columns, available = page.width - margins.left - margins.right;
    this.repeatedColumns = new Set(typeof options.wideTable === "object" ? options.wideTable.repeatColumns : []);
    let digit: number | undefined;
    if (options.columnWidths && options.columnWidths.length !== columns.length) throw new RangeError("columnWidths must declare one point width per column.");
    const minimum = typeof options.wideTable === "object" ? options.wideTable.minimumColumnWidth ?? 72 : 0;
    const widths = columns.map((column, i) => positive(options.columnWidths?.[i] ?? Math.max(minimum, column.width === undefined ? available / columns.length :
      positive(column.width, "Column width", 255) * (digit ??= fonts.select().width(48) * fontSize / 1000) + padding * 2), "PDF column width"));
    const sum = widths.reduce((n, w) => n + w, 0);
    if (sum > available + .001 && options.wideTable === "reject") throw new RangeError("Table widths exceed the printable page width.");
    this.bands = planTableBands(widths, settings);
    this.widths = !this.bands && sum > available ? widths.map(w => w * available / sum) : widths;
    let left = margins.left;
    this.lefts = this.widths.map(width => { const position = left; left += width; return position; });
    if (this.widths.some(w => w <= padding * 2)) throw new RangeError("Page is too narrow for this many columns and the requested padding.");
    this.lineHeight = fonts.select().lineHeight(fontSize);
    const depth = columns.reduce((n, c) => Math.max(n, c.groups?.length ?? 0), 0);
    if (depth > 16) throw new RangeError("Grouped headings support at most 16 levels.");
    if (depth && options.includeHeader === false) throw new TypeError("Grouped headings require leaf headings.");
    const headings: HeadingCell[][] = [];
    if (options.includeHeader !== false) {
      for (let level = 0; level < depth; level++) {
        const cells: HeadingCell[] = [];
        for (let first = 0; first < columns.length;) {
          let last = first;
          const group = columns[first]!.groups?.[level] ?? "", prefix = JSON.stringify(columns[first]!.groups?.slice(0, level + 1));
          while (group && last + 1 < columns.length && JSON.stringify(columns[last + 1]!.groups?.slice(0, level + 1)) === prefix) last++;
          cells.push({ first, span: last - first + 1, text: group }); first = last + 1;
        }
        headings.push(cells);
      }
      headings.push(columns.map((c, first) => ({ first, span: 1, text: c.header })));
    }
    this.headings = headings;
    this.totals = createTotals(columns, options.footer?.totals);
  }
  private async cell(value: ExportValue, first: number, span: number, row: number, style: CellPresentation): Promise<PdfLayoutCell> {
    const { options, padding, fontSize, budget } = this.settings, column = options.columns[first]!;
    const own = value instanceof ExportCell ? value.presentation : undefined;
    const link = value instanceof ExportCell && value.link ? copyExportLink(value.link) : undefined;
    const combined = presentation({ ...(column.alignment === undefined ? {} : { alignment: column.alignment }),
      ...(column.wrapText === undefined ? {} : { wrapText: column.wrapText }), ...style, ...own });
    const font = this.fonts.select(combined.bold, combined.italic), width = this.widths.slice(first, first + span).reduce((n, w) => n + w, 0);
    const context = { rowIndex: row, columnIndex: first, column }, raw = value instanceof ExportCell ? value.value : value;
    let text: PdfText;
    if (options.formatRuns && row >= 0 && !(value instanceof ExportCell && value.text !== undefined)) {
      const runs = options.formatRuns(raw, context);
      if (!Array.isArray(runs)) {
        if (runs && typeof (runs as { then?: unknown }).then === "function") void Promise.resolve(runs).catch(() => {});
        throw new TypeError("formatRuns must return a synchronous array of PDF text runs.");
      }
      text = runs;
    } else text = displayText(value, options, context);
    const captured = capturePdfText(text, combined, this.settings, this.fonts);
    if (options.tagged && link && !captured.text.trim()) throw new TypeError("Tagged PDF cell links require nonempty displayed text.");
    if (link && captured.runs.some(run => run.link))
      throw new TypeError("A PDF cell link cannot be combined with inline run links.");
    if (this.bands) budget.release(1, captured.text.length);
    else budget.cell(captured.text, captured.text.length);
    const lines = await wrapPdfText(captured, this.fonts, fontSize, width - padding * 2, this.settings, combined.wrapText !== false);
    return { first, span, width, font, style: combined, lineHeight: lines.reduce((height, line) => Math.max(height, line.lineHeight), font.lineHeight(fontSize)),
      ...(link ? { link } : {}), ...(this.bands ? { captured } : {}),
      lines };
  }
  private row(cells: readonly PdfLayoutCell[]): PdfLayoutRow {
    const lines = Math.max(1, ...cells.map(c => c.lines.length));
    const lineHeight = Math.max(this.lineHeight, ...cells.map(c => c.lineHeight));
    return { cells, lines, lineHeight, height: lines * lineHeight + this.settings.padding * 2 };
  }
  async headers(): Promise<readonly PdfLayoutRow[]> {
    if (this.settings.options.headerRows) return this.block(this.settings.options.headerRows, this.settings.options.headerPresentation ?? {});
    const { formatValue: _format, ...options } = this.settings.options;
    // Header labels are already resolved text and never passed through a data-value formatter.
    const rows: PdfLayoutRow[] = [];
    for (const heading of this.headings) {
      const cells: PdfLayoutCell[] = [];
      for (const c of heading) cells.push(await this.cell(new ExportCell(c.text, { text: c.text }), c.first, c.span, -1, options.headerPresentation ?? {}));
      rows.push(this.row(cells));
    }
    return rows;
  }
  async block(matrix: TableSpanRows, style: CellPresentation): Promise<readonly PdfLayoutRow[]> {
    const anchors = tableSpans(matrix, this.settings.options.columns.length);
    const { formatValue: _format, ...literalOptions } = this.settings.options;
    const cells: (PdfLayoutCell & { readonly rowSpan: number })[][] = [];
    for (const row of anchors) {
      const laidOut: (PdfLayoutCell & { readonly rowSpan: number })[] = [];
      for (const c of row) {
        const value = c.value instanceof ExportCell ? c.value : new ExportCell(c.value);
        const literal = new ExportCell(value.value, { text: displayText(value, literalOptions, { rowIndex: -1, columnIndex: c.first, column: literalOptions.columns[c.first]! }),
          ...(value.presentation ? { presentation: value.presentation } : {}), ...(value.link ? { link: value.link } : {}) });
        laidOut.push({ ...await this.cell(literal, c.first, c.span, -1, style), rowSpan: c.rowSpan });
      }
      cells.push(laidOut);
    }
    const heights = cells.map(row => Math.max(this.lineHeight + this.settings.padding * 2, ...row.filter(c => c.rowSpan === 1).map(c => c.lines.length * c.lineHeight + this.settings.padding * 2)));
    cells.forEach((row, r) => row.forEach(c => {
      const existing = heights.slice(r, r + c.rowSpan).reduce((n, h) => n + h, 0), needed = c.lines.length * c.lineHeight + this.settings.padding * 2;
      if (needed > existing) for (let i = r; i < r + c.rowSpan; i++) heights[i]! += (needed - existing) / c.rowSpan;
    }));
    return cells.map((row, r) => ({ cells: row.map(c => ({ ...c, height: heights.slice(r, r + c.rowSpan).reduce((n, h) => n + h, 0) })),
      lines: Math.max(1, Math.ceil((heights[r]! - this.settings.padding * 2) / this.lineHeight)), lineHeight: this.lineHeight, height: heights[r]!, structured: true }));
  }
  /** Repeated identifiers stay complete on every vertical fragment at the row's shared line height. */
  minimumRowLines(row: PdfLayoutRow): number {
    let lines = 1;
    if (this.bands) for (const cell of row.cells) if (this.repeatedColumns.has(cell.first)) lines = Math.max(lines, cell.lines.length);
    return lines;
  }
  minimumRowHeight(row: PdfLayoutRow): number { return this.minimumRowLines(row) * row.lineHeight + this.settings.padding * 2; }
  repeatedColumn(index: number): boolean { return this.repeatedColumns.has(index); }
  /** Measure the physical block before kept content is painted. Structured band projections are
   * shared with the painter, so measurement cannot disagree with narrower span intersections. */
  async blockHeight(rows: readonly PdfLayoutRow[]): Promise<number> {
    const physical = this.bands ? (await this.bandRows(rows, true))[0]! : rows;
    return physical.reduce((height, row) => height + row.height, 0);
  }
  /** Project already captured cells. Cache only bounded structured headings/footer matrices;
   * ordinary data windows remain owned and released by PdfTableBandFlow. */
  async bandRows(rows: readonly PdfLayoutRow[], structured = false): Promise<readonly (readonly PdfLayoutRow[])[]> {
    if (!this.bands) throw new Error("Horizontal bands were not planned.");
    if (!structured) return this.projectBandRows(rows, false);
    let prepared = this.structuredBandRows.get(rows);
    if (!prepared) { prepared = this.projectBandRows(rows, true); this.structuredBandRows.set(rows, prepared); }
    return prepared;
  }
  private async projectBandRows(rows: readonly PdfLayoutRow[], structured: boolean): Promise<readonly (readonly PdfLayoutRow[])[]> {
    const projected = [];
    for (const band of this.bands!) projected.push(await projectBandRows(rows, band, this.widths, this.settings, this.fonts));
    return structured ? alignBandBlocks(projected, this.settings) : projected;
  }
  async data(values: readonly ExportValue[], index: number, footer = false): Promise<PdfLayoutRow> {
    const { options } = this.settings;
    const cells: PdfLayoutCell[] = [];
    for (let column = 0; column < options.columns.length; column++) {
      const value = values[column], raw = value instanceof ExportCell ? value.value : value;
      if (!footer) this.totals[column]!.accept(raw);
      const style = footer ? options.footerPresentation ?? {} : index % 2 && options.alternateRowColor ? { background: options.alternateRowColor } : {};
      cells.push(await this.cell(value, column, 1, index, style));
    }
    return this.row(cells);
  }
  async footer(index: number): Promise<readonly PdfLayoutRow[] | undefined> {
    if (!this.settings.options.footer) return undefined;
    if (this.settings.options.footer.rows) return this.block(this.settings.options.footer.rows, this.settings.options.footerPresentation ?? {});
    const values = this.totals.map((total, i) => total.operation ? total.value() : this.settings.options.footer?.values?.[i]);
    return [await this.data(values, index, true)];
  }
}

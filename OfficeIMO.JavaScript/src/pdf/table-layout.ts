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
import { displayText, wrapText } from "./text.js";
import type { PdfTextLine } from "./text.js";

export interface PdfLayoutCell {
  readonly first: number;
  readonly span: number;
  readonly width: number;
  readonly font: PdfFontResource;
  readonly style: Readonly<CellPresentation>;
  readonly lines: readonly PdfTextLine[];
  readonly lineHeight: number;
  readonly rowSpan?: number;
  readonly height?: number;
  readonly link?: Readonly<ExportLink>;
}
export interface PdfLayoutRow { readonly cells: readonly PdfLayoutCell[]; readonly lines: number; readonly lineHeight: number; readonly height: number; readonly structured?: boolean; }
interface HeadingCell { readonly first: number; readonly span: number; readonly text: string; }
export class PdfTableLayout {
  readonly widths: readonly number[];
  readonly lefts: readonly number[];
  readonly lineHeight: number;
  readonly headings: readonly (readonly HeadingCell[])[];
  readonly totals: readonly NumericAggregate[];
  constructor(readonly settings: PdfSettings, readonly fonts: PdfFontResources) {
    const { options, padding, fontSize, page, margins } = settings, columns = options.columns, available = page.width - margins.left - margins.right;
    let digit: number | undefined;
    if (options.columnWidths && options.columnWidths.length !== columns.length) throw new RangeError("columnWidths must declare one point width per column.");
    const widths = columns.map((column, i) => positive(options.columnWidths?.[i] ?? (column.width === undefined ? available / columns.length :
      positive(column.width, "Column width", 255) * (digit ??= fonts.select().width(48) * fontSize / 1000) + padding * 2), "PDF column width"));
    const sum = widths.reduce((n, w) => n + w, 0);
    if (sum > available + .001 && options.wideTable === "reject") throw new RangeError("Table widths exceed the printable page width.");
    this.widths = sum > available ? widths.map(w => w * available / sum) : widths;
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
  private cell(value: ExportValue, first: number, span: number, row: number, style: CellPresentation): PdfLayoutCell {
    const { options, padding, fontSize, budget, limits } = this.settings, column = options.columns[first]!;
    const own = value instanceof ExportCell ? value.presentation : undefined;
    const combined = presentation({ ...(column.alignment === undefined ? {} : { alignment: column.alignment }),
      ...(column.wrapText === undefined ? {} : { wrapText: column.wrapText }), ...style, ...own });
    const font = this.fonts.select(combined.bold, combined.italic), width = this.widths.slice(first, first + span).reduce((n, w) => n + w, 0);
    const text = displayText(value, options, { rowIndex: row, columnIndex: first, column });
    budget.cell(text);
    return { first, span, width, font, style: combined, lineHeight: font.lineHeight(fontSize),
      ...(value instanceof ExportCell && value.link ? { link: copyExportLink(value.link) } : {}),
      lines: wrapText(text, font, fontSize, width - padding * 2, limits.maxCellCharacters, limits.maxRowLines, combined.wrapText !== false) };
  }
  private row(cells: readonly PdfLayoutCell[]): PdfLayoutRow {
    const lines = Math.max(1, ...cells.map(c => c.lines.length));
    const lineHeight = Math.max(this.lineHeight, ...cells.map(c => c.lineHeight));
    return { cells, lines, lineHeight, height: lines * lineHeight + this.settings.padding * 2 };
  }
  headers(): readonly PdfLayoutRow[] {
    if (this.settings.options.headerRows) return this.block(this.settings.options.headerRows, this.settings.options.headerPresentation ?? {});
    const { formatValue: _format, ...options } = this.settings.options;
    // Header labels are already resolved text and never passed through a data-value formatter.
    return this.headings.map(cells => this.row(cells.map(c => this.cell(new ExportCell(c.text, { text: c.text }), c.first, c.span, -1, options.headerPresentation ?? {}))));
  }
  block(matrix: TableSpanRows, style: CellPresentation): readonly PdfLayoutRow[] {
    const anchors = tableSpans(matrix, this.settings.options.columns.length);
    const { formatValue: _format, ...literalOptions } = this.settings.options;
    const cells = anchors.map(row => row.map(c => {
      const value = c.value instanceof ExportCell ? c.value : new ExportCell(c.value);
      const literal = new ExportCell(value.value, { text: displayText(value, literalOptions, { rowIndex: -1, columnIndex: c.first, column: literalOptions.columns[c.first]! }),
        ...(value.presentation ? { presentation: value.presentation } : {}), ...(value.link ? { link: value.link } : {}) });
      return { ...this.cell(literal, c.first, c.span, -1, style), rowSpan: c.rowSpan };
    }));
    const heights = cells.map(row => Math.max(this.lineHeight + this.settings.padding * 2, ...row.filter(c => c.rowSpan === 1).map(c => c.lines.length * c.lineHeight + this.settings.padding * 2)));
    cells.forEach((row, r) => row.forEach(c => {
      const existing = heights.slice(r, r + c.rowSpan).reduce((n, h) => n + h, 0), needed = c.lines.length * c.lineHeight + this.settings.padding * 2;
      if (needed > existing) for (let i = r; i < r + c.rowSpan; i++) heights[i]! += (needed - existing) / c.rowSpan;
    }));
    return cells.map((row, r) => ({ cells: row.map(c => ({ ...c, height: heights.slice(r, r + c.rowSpan).reduce((n, h) => n + h, 0) })),
      lines: Math.max(1, Math.ceil((heights[r]! - this.settings.padding * 2) / this.lineHeight)), lineHeight: this.lineHeight, height: heights[r]!, structured: true }));
  }
  data(values: readonly ExportValue[], index: number, footer = false): PdfLayoutRow {
    const { options } = this.settings;
    return this.row(options.columns.map((_, column) => {
      const value = values[column], raw = value instanceof ExportCell ? value.value : value;
      if (!footer) this.totals[column]!.accept(raw);
      const style = footer ? options.footerPresentation ?? {} : index % 2 && options.alternateRowColor ? { background: options.alternateRowColor } : {};
      return this.cell(value, column, 1, index, style);
    }));
  }
  footer(index: number): readonly PdfLayoutRow[] | undefined {
    if (!this.settings.options.footer) return undefined;
    if (this.settings.options.footer.rows) return this.block(this.settings.options.footer.rows, this.settings.options.footerPresentation ?? {});
    const values = this.totals.map((total, i) => total.operation ? total.value() : this.settings.options.footer?.values?.[i]);
    return [this.data(values, index, true)];
  }
}

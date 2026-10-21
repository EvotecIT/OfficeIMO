import { OfficeIMOError } from "../core/errors.js";
import { checkAbort } from "../core/iteration.js";
import type { PdfLayoutRow, PdfTableLayout } from "./table-layout.js";
import type { PdfPreparedParagraph } from "./paragraph-layout.js";

/** @internal A window entry is one aligned logical row fragment or an indivisible span block. */
interface BandWindowEntry { readonly rows: readonly (readonly PdfLayoutRow[])[]; readonly height: number; readonly count: number; }
interface BandWindowPainter {
  readonly firstCapacity: number;
  readonly firstFullCapacity: number;
  moveFirstToFreshPage(): Promise<void>;
  readonly continuationCapacity: number;
  paint(band: number, headers: readonly PdfLayoutRow[], rows: readonly PdfLayoutRow[], caption: PdfPreparedParagraph | undefined,
    firstWindow: boolean, firstBand: boolean): Promise<void>;
}

/** One source pass and one bounded window. A slow sink is awaited before accepting further source rows.
 * Data projection, formatting and numeric totals belong to the logical layout, never the physical bands. */
export class PdfTableBandFlow {
  private entries: BandWindowEntry[] = [];
  private height = 0;
  private count = 0;
  private firstWindow = true;
  private firstCapacity: number;
  private painted = false;
  constructor(private readonly layout: PdfTableLayout, private readonly headings: readonly (readonly PdfLayoutRow[])[],
    private readonly painter: BandWindowPainter, private readonly caption?: PdfPreparedParagraph, private readonly continuation?: PdfPreparedParagraph) { this.firstCapacity = painter.firstCapacity; }
  private get capacity(): number { return this.firstWindow ? this.firstCapacity : this.painter.continuationCapacity; }
  private async accept(entry: BandWindowEntry): Promise<void> {
    if (this.entries.length && (this.height + entry.height > this.capacity + .001 || this.count + entry.count > this.layout.settings.limits.maxTableWindowRows)) await this.flush();
    if (this.firstWindow && !this.entries.length && entry.height > this.capacity + .001 && entry.height <= this.painter.firstFullCapacity + .001) {
      await this.painter.moveFirstToFreshPage(); this.firstCapacity = this.painter.firstFullCapacity;
    }
    if (entry.height > this.capacity + .001) throw new RangeError("PDF band window cannot fit the requested table fragment.");
    this.entries.push(entry); this.height += entry.height; this.count += entry.count;
    if (this.count >= this.layout.settings.limits.maxTableWindowRows) await this.flush();
  }
  async row(row: PdfLayoutRow): Promise<void> {
    const { padding, options } = this.layout.settings, minimumLines = this.layout.minimumRowLines(row), minimumHeight = this.layout.minimumRowHeight(row);
    if (minimumHeight > this.painter.firstFullCapacity + .001 || minimumHeight > this.painter.continuationCapacity + .001)
      throw new RangeError("Repeated PDF identifiers must fit intact below the first and continued band headings/context.");
    // A normal row remains intact whenever it fits below a fresh set of repeated context/headings.
    if (this.entries.length && row.height <= this.painter.continuationCapacity + .001 && row.height > this.capacity - this.height + .001) await this.flush();
    let first = 0;
    while (first < row.lines) {
      checkAbort(options.signal);
      if (this.entries.length && this.capacity - this.height + .001 < minimumHeight) await this.flush();
      if (this.firstWindow && !this.entries.length && this.capacity + .001 < minimumHeight) {
        await this.painter.moveFirstToFreshPage(); this.firstCapacity = this.painter.firstFullCapacity;
      }
      if (this.capacity - this.height + .001 < minimumHeight)
        throw new RangeError("PDF band window cannot fit its complete repeated identifiers.");
      const lines = Math.min(row.lines - first, Math.floor((this.capacity - this.height - padding * 2 + .001) / row.lineHeight));
      if (lines < 1) throw new RangeError("PDF page cannot fit a table line below horizontal band headings.");
      const paintedLines = Math.max(lines, minimumLines), height = paintedLines * row.lineHeight + padding * 2;
      const projected = await this.layout.bandRows([row]);
      const fragments = projected.map((rows, index) => [{ ...rows[0]!, lines: paintedLines, height,
        cells: rows[0]!.cells.map(cell => {
          const source = this.layout.bands![index]!.columns[cell.first]!;
          const visible = this.layout.repeatedColumn(source) ? cell.lines : cell.lines.slice(first, first + lines);
          const { height: _height, ...plain } = cell;
          return { ...plain, lines: visible };
        }) }]);
      await this.accept({ rows: fragments, height, count: 1 }); first += lines;
      if (first < row.lines) await this.flush();
    }
  }
  async block(rows: readonly PdfLayoutRow[]): Promise<void> {
    const projected = await this.layout.bandRows(rows, true), height = projected[0]!.reduce((sum, row) => sum + row.height, 0);
    if (rows.length > this.layout.settings.limits.maxTableWindowRows) throw new OfficeIMOError("RESOURCE_LIMIT", "A structured PDF block exceeds maxTableWindowRows.");
    await this.accept({ rows: projected, height, count: rows.length });
  }
  async keptRows(rows: readonly PdfLayoutRow[]): Promise<void> {
    const height = rows.reduce((sum, row) => sum + row.height, 0);
    if (this.entries.length && height <= this.painter.continuationCapacity + .001 && height > this.capacity - this.height + .001) await this.flush();
    for (const row of rows) await this.row(row);
  }
  async flush(forceHeadings = false): Promise<void> {
    if (!this.entries.length && (!forceHeadings || this.painted || !this.headings.some(rows => rows.length))) return;
    for (let band = 0; band < this.layout.bands!.length; band++) {
      checkAbort(this.layout.settings.options.signal);
      await this.painter.paint(band, this.headings[band]!, this.entries.flatMap(entry => entry.rows[band]!),
        this.firstWindow ? this.caption : this.continuation, this.firstWindow, band === 0);
    }
    this.entries = []; this.height = 0; this.count = 0; this.firstWindow = false; this.painted = true;
  }
}

import { OfficeIMOError } from "../core/errors.js";
import type { PdfSettings } from "./settings.js";
import type { PdfLayoutCell, PdfLayoutRow } from "./table-layout.js";
import type { PdfFontResources } from "./font-resources.js";
import { wrapPdfText } from "./inline-text.js";

/** @internal One physical column order; repeated identifiers precede the next ordered data columns. */
export interface PdfTableBand { readonly columns: readonly number[]; readonly lefts: readonly number[]; }

/** Plan without opening a source. Exact column widths never shrink in band mode. */
export function planTableBands(widths: readonly number[], settings: PdfSettings): readonly PdfTableBand[] | undefined {
  const policy = settings.options.wideTable;
  if (typeof policy !== "object") return undefined;
  const available = settings.page.width - settings.margins.left - settings.margins.right, repeats = policy.repeatColumns ?? [];
  if (repeats.some(index => index >= widths.length)) throw new RangeError("repeatColumns index exceeds the declared columns.");
  const repeated = new Set(repeats), rest = widths.map((_, index) => index).filter(index => !repeated.has(index));
  const prefix = repeats.reduce((sum, index) => sum + widths[index]!, 0);
  if (prefix > available + .001) throw new RangeError("Repeated PDF columns exceed the printable width.");
  const groups: number[][] = [];
  let current = [...repeats], used = prefix;
  for (const index of rest) {
    const width = widths[index]!;
    if (prefix + width > available + .001) throw new RangeError("A PDF column and its repeated identifiers cannot fit a band.");
    if (used + width > available + .001) { groups.push(current); current = [...repeats]; used = prefix; }
    current.push(index); used += width;
  }
  if (current.length) groups.push(current);
  if (groups.length > settings.limits.maxTableBands) throw new OfficeIMOError("RESOURCE_LIMIT", "maxTableBands exceeded.");
  if (settings.limits.maxTableWindowRows < 1) throw new OfficeIMOError("RESOURCE_LIMIT", "maxTableWindowRows must allow at least one row fragment.");
  return Object.freeze(groups.map(columns => {
    let x = settings.margins.left;
    const lefts = columns.map(index => { const left = x; x += widths[index]!; return left; });
    return Object.freeze({ columns: Object.freeze(columns), lefts: Object.freeze(lefts) });
  }));
}

/** A span crossing a band boundary repeats its captured text at each visible intersection.
 * Runs/formatters are never recaptured, and vertical spans retain their original row ownership. */
export async function projectBandRows(rows: readonly PdfLayoutRow[], band: PdfTableBand, widths: readonly number[],
  settings: PdfSettings, fonts: PdfFontResources): Promise<readonly PdfLayoutRow[]> {
  const result: PdfLayoutRow[] = [];
  for (const row of rows) {
    const cells: PdfLayoutCell[] = [];
    for (const cell of row.cells) {
      for (let first = 0; first < band.columns.length;) {
        const source = band.columns[first]!;
        if (source < cell.first || source >= cell.first + cell.span) { first++; continue; }
        let end = first + 1;
        while (end < band.columns.length && band.columns[end] === band.columns[end - 1]! + 1 && band.columns[end]! < cell.first + cell.span) end++;
        const width = band.columns.slice(first, end).reduce((sum, index) => sum + widths[index]!, 0);
        const lines = cell.span === 1 ? cell.lines : await wrapPdfText(cell.captured!, fonts, settings.fontSize, width - settings.padding * 2, settings, cell.style.wrapText !== false);
        cells.push({ ...cell, source: cell.source ?? cell, first, span: end - first, width, lines,
          lineHeight: lines.reduce((height, line) => Math.max(height, line.lineHeight), cell.font.lineHeight(settings.fontSize)) });
        first = end;
      }
    }
    cells.sort((a, b) => a.first - b.first);
    result.push({ ...row, source: row.source ?? row, cells });
  }
  return result;
}

/** Equal row heights across every physical band keep structured headings and row windows aligned. */
export function alignBandBlocks(bands: readonly (readonly PdfLayoutRow[])[], settings: PdfSettings): readonly (readonly PdfLayoutRow[])[] {
  if (!bands.length || !bands[0]!.length) return bands;
  const rows = bands[0]!.length;
  const heights = Array.from({ length: rows }, (_, index) => Math.max(...bands.map(band => Math.max(settings.padding * 2 + band[index]!.lineHeight,
    ...band[index]!.cells.filter(cell => (cell.rowSpan ?? 1) === 1).map(cell => cell.lines.length * cell.lineHeight + settings.padding * 2)))));
  for (const band of bands) for (const [index, row] of band.entries()) for (const cell of row.cells) {
    const span = cell.rowSpan ?? 1, height = heights.slice(index, index + span).reduce((sum, value) => sum + value, 0), needed = cell.lines.length * cell.lineHeight + settings.padding * 2;
    if (needed > height) for (let i = index; i < index + span; i++) heights[i]! += (needed - height) / span;
  }
  return bands.map(band => band.map((row, index) => ({ ...row, structured: true, height: heights[index]!,
    cells: row.cells.map(cell => ({ ...cell, height: heights.slice(index, index + (cell.rowSpan ?? 1)).reduce((sum, value) => sum + value, 0) })) })));
}

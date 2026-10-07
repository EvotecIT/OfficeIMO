import { OfficeIMOError } from "../core/errors.js";
import { cellPosition } from "./attachments.js";

interface Region { readonly ref: string; readonly top: number; readonly bottom: number; readonly left: number; readonly right: number; }
interface RowIndex { readonly middle: number; readonly starts: readonly Region[]; readonly ends: readonly Region[]; readonly before?: RowIndex; readonly after?: RowIndex; }
function indexRows(regions: readonly Region[]): RowIndex | undefined {
  if (!regions.length) return undefined;
  const ordered = [...regions].sort((a, b) => a.top - b.top), middle = ordered[Math.floor(ordered.length / 2)]!.top;
  const crossing: Region[] = [], before: Region[] = [], after: Region[] = [];
  for (const region of ordered) (region.bottom < middle ? before : region.top > middle ? after : crossing).push(region);
  const left = indexRows(before), right = indexRows(after);
  return { middle, starts: crossing, ends: [...crossing].sort((a, b) => b.bottom - a.bottom), ...(left ? { before: left } : {}), ...(right ? { after: right } : {}) };
}
function rowRegions(index: RowIndex | undefined, row: number, output: Region[]): void {
  if (!index) return;
  for (const region of row <= index.middle ? index.starts : index.ends) {
    if (row <= index.middle ? region.top > row : region.bottom < row) break;
    output.push(region);
  }
  rowRegions(row < index.middle ? index.before : row > index.middle ? index.after : undefined, row, output);
}
/** @internal Retained metadata only. Row lookup avoids scanning every merge for every exported cell. */
export class MergeRegions {
  readonly references: readonly string[];
  private readonly index: RowIndex | undefined;
  private currentRow = -1;
  private current: Region[] = [];
  private readonly lastRow: number;
  constructor(references: readonly string[], columns: number, maximum: number) {
    if (references.length > maximum) throw new OfficeIMOError("RESOURCE_LIMIT", "maxMergedRanges exceeded.");
    if (!references.length) { this.references = Object.freeze([]); this.lastRow = 0; this.index = undefined; return; }
    const regions = references.map(ref => {
      if (typeof ref !== "string" || ref.split(":").length !== 2) throw new TypeError("Merged ranges require uppercase A1:B2 references.");
      const [first, last] = ref.split(":"), start = cellPosition(first!), end = cellPosition(last!);
      if (end.row < start.row || end.column < start.column || (end.row === start.row && end.column === start.column)) throw new RangeError("Merged ranges require ordered distinct cells.");
      if (end.column > columns) throw new RangeError("Merged ranges must stay within declared columns.");
      return { ref, top: start.row, bottom: end.row, left: start.column, right: end.column };
    });
    // A row sweep and column range tree detect rectangle overlaps without a quadratic pair scan.
    const events = regions.flatMap(r => [{ row: r.top, delta: 1, region: r }, { row: r.bottom + 1, delta: -1, region: r }]).sort((a, b) => a.row - b.row || a.delta - b.delta);
    const counts = new Int32Array(65536), lazy = new Int32Array(65536);
    function update(node: number, first: number, last: number, left: number, right: number, delta: number): void {
      if (left <= first && last <= right) { counts[node] = counts[node]! + delta; lazy[node] = lazy[node]! + delta; return; }
      const middle = (first + last) >>> 1;
      if (left <= middle) update(node * 2, first, middle, left, right, delta);
      if (right > middle) update(node * 2 + 1, middle + 1, last, left, right, delta);
      counts[node] = lazy[node]! + Math.max(counts[node * 2]!, counts[node * 2 + 1]!);
    }
    for (const event of events) {
      update(1, 1, 16384, event.region.left, event.region.right, event.delta);
      if (counts[1]! > 1) throw new TypeError("Merged ranges must not overlap: " + event.region.ref);
    }
    this.references = Object.freeze(regions.map(r => r.ref)); this.lastRow = regions.reduce((last, r) => Math.max(last, r.bottom), 0); this.index = indexRows(regions);
  }
  covered(column: number, row: number): boolean {
    if (!this.index || row > this.lastRow) return false;
    if (this.currentRow !== row) { this.current = []; rowRegions(this.index, row, this.current); this.current.sort((a, b) => a.left - b.left); this.currentRow = row; }
    let first = 0, last = this.current.length - 1;
    while (first <= last) {
      const middle = (first + last) >>> 1, region = this.current[middle]!;
      if (column < region.left) last = middle - 1;
      else if (column > region.right) first = middle + 1;
      else return row !== region.top || column !== region.left;
    }
    return false;
  }
  validateRows(rows: number): void { if (this.lastRow > rows) throw new RangeError("Merged ranges must stay within exported rows."); }
}

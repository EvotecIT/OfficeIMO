import type { TableSpanRows } from "../core/index.js";
import { ExportCell } from "../core/presentation.js";
import { tableSpans } from "../internal/table-spans.js";
import { copyValue, columnName } from "./values.js";

/** Bounded heading/footer values and relative merges captured before row production. */
export class StructuredRows {
  readonly values: readonly (readonly unknown[])[];
  private readonly ranges: readonly { first: number; last: number; top: number; bottom: number }[];
  constructor(rows: TableSpanRows, columns: number) {
    const anchors = tableSpans(rows, columns), ranges: { first: number; last: number; top: number; bottom: number }[] = [];
    this.values = anchors.map((row, level) => {
      const values: unknown[] = Array(columns).fill(undefined);
      for (const anchor of row) {
        values[anchor.first] = copyValue(anchor.value);
        if (anchor.span > 1 || anchor.rowSpan > 1) ranges.push({ first: anchor.first + 1, last: anchor.first + anchor.span, top: level + 1, bottom: level + anchor.rowSpan });
      }
      return Object.freeze(values);
    });
    this.ranges = ranges;
  }
  get mergeCount(): number { return this.ranges.length; }
  merges(offset: number): readonly string[] { return this.ranges.map(r => columnName(r.first) + (r.top + offset) + ":" + columnName(r.last) + (r.bottom + offset)); }
  /** A native filter/table requires one real leaf cell for each declared column. */
  matchesLeaf(headers: readonly string[]): boolean {
    const last = this.values.length;
    return !this.ranges.some(r => r.bottom === last) && this.values[last - 1]!.every((value, index) => (value instanceof ExportCell ? value.value : value) === headers[index]);
  }
}

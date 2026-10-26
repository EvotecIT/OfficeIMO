import type { ColumnOutline, RowOutline, SheetOptions } from "./types.js";
import { captureFields } from "../internal/capture.js";

/** Capture bounded disjoint ranges before asynchronous serialization; absent columns stay visible. */
export function captureColumnOutlines(input: readonly ColumnOutline[] | undefined, columns: number): readonly ColumnOutline[] | undefined {
  if (input === undefined) return undefined;
  if (!Array.isArray(input) || input.length > columns) throw new TypeError("Column outlines must be bounded, non-overlapping ranges.");
  let end = 0;
  return Object.freeze(input.map(range => {
    if (!range || typeof range !== "object") throw new TypeError("Column outlines require range metadata.");
    return captureFields(range, ["firstColumn", "columnCount", "level", "hidden", "collapsed"]);
  }).sort((a, b) => a.firstColumn - b.firstColumn).map(range => {
    const { firstColumn, columnCount, level = 0, hidden, collapsed } = range;
    if (!Number.isSafeInteger(firstColumn) || firstColumn < end || !Number.isSafeInteger(columnCount) || columnCount < 1 || firstColumn + columnCount > columns)
      throw new RangeError("Column outline ranges must be disjoint and within the declared columns.");
    if (!Number.isInteger(level) || level < 0 || level > 7) throw new RangeError("Column outline level must be from zero through seven.");
    for (const flag of [hidden, collapsed]) if (flag !== undefined && typeof flag !== "boolean") throw new TypeError("Column outline flags must be booleans.");
    end = firstColumn + columnCount;
    return Object.freeze({ firstColumn, columnCount, level, ...(hidden === undefined ? {} : { hidden }), ...(collapsed === undefined ? {} : { collapsed }) });
  }));
}

/** @internal Positions are captured even when a sheet has only column outlines. */
export function validateOutlinePositions(options: SheetOptions): void {
  for (const flag of [options.outlineSummaryBelow, options.outlineSummaryRight])
    if (flag !== undefined && typeof flag !== "boolean") throw new TypeError("Outline summary positions must be booleans.");
}

/** @internal Resolve at most 16,384 columns once, without retaining any worksheet body. */
export function columnOutlineMap(ranges: readonly ColumnOutline[] | undefined): ReadonlyMap<number, RowOutline> {
  const result = new Map<number, RowOutline>();
  for (const range of ranges ?? []) for (let index = range.firstColumn; index < range.firstColumn + range.columnCount; index++) result.set(index, range);
  return result;
}

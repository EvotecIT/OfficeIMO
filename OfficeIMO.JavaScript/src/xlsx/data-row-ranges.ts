import { OfficeIMOError } from "../core/errors.js";
import { columnName } from "./values.js";

/** One-based physical worksheet rows, excluding generated headings and summaries. */
export interface DataRowRange { readonly firstRow: number; readonly lastRow: number; }

const scopes = new WeakMap<object, readonly DataRowRange[]>();

/** @internal A format adapter owns this bounded list until worksheet completion. */
export function retainDataRowRanges(worksheet: object, ranges: readonly DataRowRange[]): void { scopes.set(worksheet, ranges); }

/** @internal Ordinary worksheets retain their contiguous data-column contract. */
export function dataRowRanges(worksheet: object): readonly DataRowRange[] | undefined { return scopes.get(worksheet); }

/** @internal Count accepted details while preserving physical row counts for native references. */
export function dataRowCount(worksheet: object, headerRows: number, serializedRows: number): number {
  const ranges = scopes.get(worksheet);
  if (ranges === undefined) return serializedRows;
  const lastSerialized = headerRows + serializedRows;
  let rows = 0;
  for (const range of ranges) {
    if (range.firstRow > lastSerialized) break;
    rows += Math.min(range.lastRow, lastSerialized) - range.firstRow + 1;
  }
  return rows;
}

/** @internal Native editing rules share one finite, bounded space-separated reference. */
export function dataRowReference(firstColumn: number, lastColumn: number, headerRows: number, dataRows: number,
  ranges?: readonly DataRowRange[]): string | undefined {
  if (ranges === undefined) return dataRows ? columnName(firstColumn) + (headerRows + 1) + ":" + columnName(lastColumn) + (headerRows + dataRows) : undefined;
  const first = columnName(firstColumn), last = columnName(lastColumn);
  let reference = "", previous = headerRows;
  for (const range of ranges) {
    if (!Number.isSafeInteger(range.firstRow) || !Number.isSafeInteger(range.lastRow) || range.firstRow <= previous ||
      range.lastRow < range.firstRow || range.lastRow > headerRows + dataRows)
      throw new RangeError("Detail editing ranges must be ordered within serialized worksheet rows.");
    const next = first + range.firstRow + ":" + last + range.lastRow;
    if (reference.length + (reference ? 1 : 0) + next.length > 32767)
      throw new OfficeIMOError("RESOURCE_LIMIT", "Detail editing references exceed 32,767 UTF-16 characters; split the worksheet.");
    reference += (reference ? " " : "") + next; previous = range.lastRow;
  }
  return reference || undefined;
}

import type { ProjectionColumn } from "../internal/columns.js";
import { columnName } from "./values.js";

/** @internal Data-column selectors use stable keys or one-based worksheet columns. */
export function resolveDataColumn(columns: readonly ProjectionColumn[], value: string | number, context = "Data"): number {
  if (typeof value === "string") {
    const matches = columns.flatMap((column, index) => column.key === value ? [index + 1] : []);
    if (matches.length !== 1) throw new TypeError(context + " column keys must identify exactly one declared column.");
    return matches[0]!;
  }
  if (!Number.isInteger(value) || value < 1 || value > columns.length)
    throw new RangeError(context + " column numbers must be one-based and within declared columns.");
  return value;
}

/** @internal Formula references use literal final worksheet names and finite physical coordinates. */
export function qualifiedColumnRange(sheetName: string, column: number, firstRow: number, rows: number): string {
  if (!Number.isSafeInteger(firstRow) || firstRow < 1 || !Number.isSafeInteger(rows) || rows < 1 || firstRow + rows - 1 > 1048576)
    throw new RangeError("A data range must fit within Excel's worksheet rows.");
  const letter = columnName(column);
  return "'" + sheetName.replace(/'/g, "''") + "'!$" + letter + "$" + firstRow + ":$" + letter + "$" + (firstRow + rows - 1);
}

/** @internal A series header is a single cell rather than a one-cell range. */
export function qualifiedCell(sheetName: string, column: number, row: number): string {
  return qualifiedColumnRange(sheetName, column, row, 1).split(":")[0]!;
}

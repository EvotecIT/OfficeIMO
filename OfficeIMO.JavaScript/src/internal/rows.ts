import type { Column } from "../core/index.js";

export function rowValues(row: unknown, columns: readonly Column[]): readonly unknown[] {
  if (Array.isArray(row)) {
    if (row.length > columns.length) throw new RangeError("Row has more values than declared columns.");
    return row;
  }
  if (!row || typeof row !== "object" || row instanceof Date) throw new TypeError("A row must be an array or object.");
  return columns.map(c => (row as Record<string, unknown>)[c.key ?? c.header]);
}

export function copyColumns(columns: readonly Column[]): Column[] {
  if (!Array.isArray(columns)) throw new TypeError("Declare the columns in export order.");
  return columns.map(c => {
    if (!c || typeof c.header !== "string" || (c.key !== undefined && typeof c.key !== "string"))
      throw new TypeError("Each column needs a string header and an optional string key.");
    return { ...c };
  });
}

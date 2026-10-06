import type { Column } from "../core/index.js";

export function rowValues(row: unknown, columns: readonly Column[]): readonly unknown[] {
  if (Array.isArray(row)) {
    if (row.length > columns.length) throw new RangeError("Row has more values than declared columns.");
    return row;
  }
  if (!row || typeof row !== "object" || row instanceof Date) throw new TypeError("A row must be an array or object.");
  return columns.map(c => {
    const key = c.key ?? c.header;
    return Object.prototype.hasOwnProperty.call(row, key) ? (row as Record<string, unknown>)[key] : undefined;
  });
}

export function copyColumns(columns: readonly Column[]): Column[] {
  if (!Array.isArray(columns)) throw new TypeError("Declare the columns in export order.");
  return columns.map(c => {
    if (!c || typeof c.header !== "string" || (c.key !== undefined && typeof c.key !== "string"))
      throw new TypeError("Each column needs a string header and an optional string key.");
    if (c.groups !== undefined && (!Array.isArray(c.groups) || c.groups.some((group: unknown) => typeof group !== "string"))) throw new TypeError("Column groups must be an array of strings.");
    return { ...c, ...(c.groups ? { groups: Object.freeze([...c.groups]) } : {}) };
  });
}

import type { ColumnValueContext } from "../core/index.js";
import type { ProjectionColumn } from "./columns.js";
import { assertExportValue } from "../core/presentation.js";
import { checkAbort } from "../core/iteration.js";

export function createRowProjector(columns: readonly ProjectionColumn[],
  worksheet?: { readonly sheetName: string; readonly firstDataRow: number }, signal?: AbortSignal,
  validate: (value: unknown) => void = assertExportValue): (row: unknown, rowIndex?: number) => readonly unknown[] {
  const getters = columns.some(column => column.value);
  return (row, rowIndex = 0) => {
    if (Array.isArray(row) && !getters) {
      if (row.length > columns.length) throw new RangeError("Row has more values than declared columns.");
      for (const value of row) validate(value);
      return row;
    }
    if (!row || typeof row !== "object" || row instanceof Date) throw new TypeError("A row must be an array or object.");
    if (Array.isArray(row) && row.length > columns.length && !columns.every(column => column.value))
      throw new RangeError("Project every column explicitly when selecting from a wider array row.");
    return columns.map((c, columnIndex) => {
      if (c.value) {
        checkAbort(signal);
        const context: ColumnValueContext = { rowIndex, columnIndex, column: c,
          ...(worksheet ? { sheetName: worksheet.sheetName, worksheetRow: worksheet.firstDataRow + rowIndex } : {}) };
        const result = c.value(row as never, context);
        validate(result);
        checkAbort(signal);
        return result;
      }
      if (Array.isArray(row)) { const result = row[columnIndex]; validate(result); return result; }
      const key = c.key ?? c.header;
      const result = Object.prototype.hasOwnProperty.call(row, key) ? (row as Record<string, unknown>)[key] : undefined;
      validate(result); return result;
    });
  };
}

export function copyColumns<C extends ProjectionColumn>(columns: readonly C[], workbookStyles = false): C[] {
  if (!Array.isArray(columns)) throw new TypeError("Declare the columns in export order.");
  return columns.map(c => {
    if (!workbookStyles && (c as { style?: unknown })?.style !== undefined)
      throw new TypeError("Workbook-local column styles require the advanced Workbook API; use portable ExportCell presentation.");
    if (!c || typeof c.header !== "string" || (c.key !== undefined && typeof c.key !== "string"))
      throw new TypeError("Each column needs a string header and an optional string key.");
    if (c.value !== undefined && typeof c.value !== "function") throw new TypeError("Column value getters must be functions.");
    if (c.groups !== undefined && (!Array.isArray(c.groups) || c.groups.some((group: unknown) => typeof group !== "string"))) throw new TypeError("Column groups must be an array of strings.");
    return { ...c, ...(c.groups ? { groups: Object.freeze([...c.groups]) } : {}) };
  });
}

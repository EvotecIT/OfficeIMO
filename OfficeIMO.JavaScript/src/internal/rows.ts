import type { ColumnValueContext } from "../core/index.js";
import type { ProjectionColumn } from "./columns.js";
import { assertExportValue } from "../core/presentation.js";
import { checkAbort } from "../core/iteration.js";
import { isDate } from "../core/dates.js";
import { captureFields } from "./capture.js";

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
    if (!row || typeof row !== "object" || isDate(row)) throw new TypeError("A row must be an array or object.");
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

/** @internal Shared metadata fields for full columns and adapter-owned partial overrides. */
export const columnOptionFields = ["header", "key", "value", "width", "type", "format", "wrapText", "alignment", "groups", "style", "valueFormatter"] as const;

export function copyColumns<C extends ProjectionColumn>(columns: readonly C[], workbookStyles = false): C[] {
  if (!Array.isArray(columns)) throw new TypeError("Declare the columns in export order.");
  return columns.map(c => {
    if (!c || typeof c !== "object" && typeof c !== "function") throw new TypeError("Each column needs a string header and an optional string key.");
    // Own extension metadata remains available to its format owner. Capture the
    // supported structural fields even when a class exposes prototype getters.
    const captured = captureFields(c, columnOptionFields as readonly (keyof C)[]);
    if (!workbookStyles && (captured as { style?: unknown }).style !== undefined)
      throw new TypeError("Workbook-local column styles require the advanced Workbook API; use portable ExportCell presentation.");
    if (typeof captured.header !== "string" || (captured.key !== undefined && typeof captured.key !== "string"))
      throw new TypeError("Each column needs a string header and an optional string key.");
    if (captured.value !== undefined && typeof captured.value !== "function") throw new TypeError("Column value getters must be functions.");
    if (captured.groups !== undefined) {
      if (!Array.isArray(captured.groups)) throw new TypeError("Column groups must be an array of strings.");
      const groups: readonly unknown[] = Array.from(captured.groups);
      if (groups.some(group => typeof group !== "string")) throw new TypeError("Column groups must be an array of strings.");
      return { ...captured, groups: Object.freeze(groups) as readonly string[] };
    }
    return captured;
  });
}

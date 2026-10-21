import { ExportCell, assertExportValue } from "../core/presentation.js";
import type { ExportValue } from "../core/presentation.js";
import { isDate } from "../core/dates.js";

/** Copy portable values when an operation retains metadata or a projected row beyond its producer's turn. */
export function copyExportValue(value: unknown): ExportValue {
  assertExportValue(value);
  if (isDate(value)) return new Date(Date.prototype.getTime.call(value));
  if (value instanceof ExportCell) {
    const { value: scalar, text, presentation, link } = value;
    return new ExportCell(isDate(scalar) ? new Date(Date.prototype.getTime.call(scalar)) : scalar,
      { ...(text === undefined ? {} : { text }), ...(presentation === undefined ? {} : { presentation }),
        ...(link === undefined ? {} : { link }) });
  }
  return value;
}

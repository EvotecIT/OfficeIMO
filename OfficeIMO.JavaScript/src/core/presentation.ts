import type { Alignment, CellValue } from "./index.js";
import { isDate } from "./dates.js";
import { copyExportLink } from "./links.js";
import type { ExportLink } from "./links.js";
const exportCellBrand = Symbol.for("@evotecit/officeimo/ExportCell");

/** Portable presentation selected by the report producer, without workbook-local indexes. */
export interface CellPresentation {
  readonly background?: string;
  readonly color?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly wrapText?: boolean;
  readonly alignment?: Alignment;
  readonly numberFormat?: string;
}
export interface ExportCellOptions {
  /** Display text for PDF, CSV's display mode and bounded width sampling; typed values remain authoritative. */
  readonly text?: string;
  readonly presentation?: CellPresentation;
  /** XLSX/PDF retain the external link. CSV emits ordinary value/display text. */
  readonly link?: ExportLink;
}
/** @internal Capture only portable scalar presentation fields before retaining or cloning a cell. */
export function copyCellPresentation(value: CellPresentation): Readonly<CellPresentation> {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError("Cell presentation must be an object.");
  const result: Record<string, unknown> = {};
  for (const key of ["background", "color", "numberFormat", "bold", "italic", "wrapText", "alignment"] as const) {
    const field = value[key];
    if (field === undefined) continue;
    if (["bold", "italic", "wrapText"].includes(key)) {
      if (typeof field !== "boolean") throw new TypeError("Cell presentation " + key + " must be boolean.");
    } else if (typeof field !== "string") throw new TypeError("Cell presentation " + key + " must be text.");
    if (key === "alignment" && !["left", "center", "right", "fill", "justify", "distributed"].includes(field as string))
      throw new TypeError("Invalid cell presentation alignment.");
    result[key] = field;
  }
  return Object.freeze(result);
}
/** One resolved value/presentation decision that can be reused across exports. Strings remain literal data. */
export class ExportCell {
  static [Symbol.hasInstance](value: unknown): boolean { return !!value && typeof value === "object" && (value as Record<symbol, unknown>)[exportCellBrand] === true; }
  readonly text: string | undefined;
  readonly presentation: Readonly<CellPresentation> | undefined;
  readonly link: Readonly<ExportLink> | undefined;
  constructor(readonly value: CellValue, options: ExportCellOptions = {}) {
    const { text, presentation, link } = options;
    if (text !== undefined && typeof text !== "string") throw new TypeError("Display text must be a string.");
    assertScalar(value);
    this.text = text;
    this.presentation = presentation === undefined ? undefined : copyCellPresentation(presentation);
    this.link = link === undefined ? undefined : copyExportLink(link);
    Object.defineProperty(this, exportCellBrand, { value: true });
    Object.freeze(this);
  }
}
/** @internal Reject async formatters while observing their rejection immediately. */
export function assertScalar(value: unknown): asserts value is CellValue {
  const kind = typeof value;
  if (value == null || kind === "string" || kind === "number" || kind === "boolean" || isDate(value)) return;
  if (typeof (value as { then?: unknown }).then === "function") void Promise.resolve(value).catch(() => {});
  throw new TypeError("Export values and formatter results must be synchronous strings, numbers, booleans, Dates or null.");
}
export type ExportValue = CellValue | ExportCell;
/** @internal Validate selected values before a destination interprets presentation. */
export function assertExportValue(value: unknown): asserts value is ExportValue {
  if (value !== null && typeof value === "object" && value instanceof ExportCell) return;
  assertScalar(value);
}

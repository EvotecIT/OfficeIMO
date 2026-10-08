import type { CellPresentation } from "../core/index.js";
import { OfficeIMOError } from "../core/errors.js";
import { ExportBudget } from "../core/limits.js";
import { copyColumns } from "../internal/rows.js";
import { tableSpans } from "../internal/table-spans.js";
import type { TableSpanRows } from "../core/index.js";
import type { PdfOptions, PdfMargins, PdfPageSize, PdfLimits } from "./types.js";

export function positive(value: number, name: string, max = 14400): number {
  if (!Number.isFinite(value) || value <= 0 || value > max) throw new RangeError(name + " must be positive and at most " + max + ".");
  return value;
}
export function color(value: string): string {
  if (typeof value !== "string" || !/^#?[0-9a-f]{6}$/i.test(value)) throw new TypeError("PDF colors must be six-digit RGB hex strings.");
  const hex = value.replace(/^#/, "");
  return [0, 2, 4].map(i => String(Math.round(parseInt(hex.slice(i, i + 2), 16) / 255 * 10000) / 10000)).join(" ");
}
export function presentation(value: CellPresentation = {}): Readonly<CellPresentation> {
  if (!value || typeof value !== "object") throw new TypeError("PDF presentation must be an object.");
  if (value.background !== undefined) color(value.background);
  if (value.color !== undefined) color(value.color);
  for (const key of ["bold", "italic", "wrapText"] as const) if (value[key] !== undefined && typeof value[key] !== "boolean") throw new TypeError("PDF presentation " + key + " must be boolean.");
  if (value.alignment !== undefined && !["left", "center", "right"].includes(value.alignment)) throw new TypeError("PDF table alignment supports left, center and right.");
  return Object.freeze({ ...value });
}
const sizes: Readonly<Record<string, PdfPageSize>> = {
  A3: { width: 841.89, height: 1190.551 }, A4: { width: 595.276, height: 841.89 }, A5: { width: 419.528, height: 595.276 },
  LETTER: { width: 612, height: 792 }, LEGAL: { width: 612, height: 1008 }, TABLOID: { width: 792, height: 1224 }
};
export interface PdfSettings {
  readonly options: PdfOptions;
  readonly page: PdfPageSize;
  readonly margins: PdfMargins;
  readonly fontSize: number;
  readonly padding: number;
  readonly limits: Required<Pick<PdfLimits, "maxPages" | "maxColumns" | "maxCellCharacters" | "maxRowLines" | "maxFontBytes" | "maxPageBytes">>;
  readonly budget: ExportBudget;
}
export function settings(configuration: PdfOptions): PdfSettings {
  const columns = copyColumns(configuration?.columns).map(c => Object.freeze(c));
  if (!columns.length) throw new TypeError("PDF tables need at least one declared column.");
  function copySpans(rows: TableSpanRows): TableSpanRows {
    tableSpans(rows, columns.length);
    return Object.freeze(rows.map(row => Object.freeze(row.map(cell => cell ? Object.freeze({ ...cell }) : null))));
  }
  const options: PdfOptions = { ...configuration, columns,
    ...(configuration.headerRows ? { headerRows: copySpans(configuration.headerRows) } : {}),
    ...(configuration.fonts ? { fonts: Object.freeze({ ...configuration.fonts }) } : {}),
    ...(configuration.columnWidths ? { columnWidths: Object.freeze([...configuration.columnWidths]) } : {}),
    ...(configuration.footer ? { footer: Object.freeze({ ...configuration.footer,
      ...(configuration.footer.rows ? { rows: copySpans(configuration.footer.rows) } : {}),
      ...(configuration.footer.values ? { values: Object.freeze([...configuration.footer.values]) } : {}),
      ...(configuration.footer.totals ? { totals: Object.freeze({ ...configuration.footer.totals }) } : {}) }) } : {}),
    headerPresentation: presentation({ background: "e7edf5", bold: true, ...configuration.headerPresentation }),
    footerPresentation: presentation({ background: "eef2f6", bold: true, ...configuration.footerPresentation }) };
  const budget = new ExportBudget(options.limits);
  const limits = { maxPages: 10000, maxColumns: 1024, maxCellCharacters: 1000000, maxRowLines: 100000, maxFontBytes: 16 * 1024 * 1024, maxPageBytes: 8 * 1024 * 1024 };
  for (const key of Object.keys(limits) as (keyof typeof limits)[]) {
    const override = options.limits?.[key]; if (override !== undefined) limits[key] = override;
  }
  if (columns.length > limits.maxColumns) throw new OfficeIMOError("RESOURCE_LIMIT", "maxColumns exceeded.");
  const size = typeof options.pageSize === "string" || options.pageSize === undefined ? sizes[options.pageSize ?? "A4"] : options.pageSize;
  if (!size) throw new TypeError("Unsupported PDF page size.");
  positive(size.width, "Page width"); positive(size.height, "Page height");
  if (options.orientation !== undefined && !["portrait", "landscape"].includes(options.orientation)) throw new TypeError("Invalid PDF orientation.");
  const landscape = options.orientation === "landscape", page = { width: landscape ? Math.max(size.width, size.height) : Math.min(size.width, size.height), height: landscape ? Math.min(size.width, size.height) : Math.max(size.width, size.height) };
  const margins = { top: 36, right: 36, bottom: 36, left: 36, ...(typeof options.margins === "number" ? { top: options.margins, right: options.margins, bottom: options.margins, left: options.margins } : options.margins) };
  for (const value of Object.values(margins)) if (!Number.isFinite(value) || value < 0) throw new RangeError("PDF margins must be finite and nonnegative.");
  if (margins.left + margins.right >= page.width || margins.top + margins.bottom >= page.height) throw new RangeError("PDF margins leave no content area.");
  const fontSize = positive(options.fontSize ?? 9, "Font size", 144), padding = options.padding ?? 4;
  if (!Number.isFinite(padding) || padding < 0 || padding > 72) throw new RangeError("PDF padding must be from 0 through 72 points.");
  if (options.wideTable !== undefined && !["fit", "reject"].includes(options.wideTable)) throw new TypeError("wideTable must be fit or reject.");
  for (const key of ["title", "messageTop", "messageBottom"] as const) if (options[key] !== undefined && typeof options[key] !== "string") throw new TypeError(key + " must be text.");
  for (const key of ["pageHeader", "pageFooter"] as const) if (options[key] !== undefined && typeof options[key] !== "string" && typeof options[key] !== "function") throw new TypeError(key + " must be text or a synchronous callback.");
  if (options.formatValue !== undefined && typeof options.formatValue !== "function") throw new TypeError("formatValue must be a synchronous function.");
  for (const key of ["compression", "includeHeader", "pageNumbers"] as const) if (options[key] !== undefined && typeof options[key] !== "boolean") throw new TypeError(key + " must be boolean.");
  if (options.alternateRowColor !== undefined) color(options.alternateRowColor);
  if (options.footer?.values && options.footer.values.length > columns.length) throw new RangeError("Footer has more values than declared columns.");
  if (options.footer?.rows && (options.footer.values !== undefined || options.footer.totals !== undefined)) throw new TypeError("Structured footer rows cannot be combined with values/totals.");
  if (options.headerRows && options.includeHeader === false) throw new TypeError("headerRows requires includeHeader.");
  return { options, page, margins, fontSize, padding, limits, budget };
}

import { OfficeIMOError } from "../core/errors.js";
import { ExportBudget } from "../core/limits.js";
import { copyColumns } from "../internal/rows.js";
import { tableSpans } from "../internal/table-spans.js";
import type { TableSpanRows } from "../core/index.js";
import type { PdfOptions, PdfMargins, PdfPageSize, PdfLimits } from "./types.js";
import type { PdfDocumentOptions, PdfParagraphStyle } from "./document-types.js";
import { captureParagraphStyles } from "./paragraph-styles.js";
import { positive, color, presentation } from "./presentation-settings.js";
export { positive, color, presentation } from "./presentation-settings.js";
const sizes: Readonly<Record<string, PdfPageSize>> = {
  A3: { width: 841.89, height: 1190.551 }, A4: { width: 595.276, height: 841.89 }, A5: { width: 419.528, height: 595.276 },
  LETTER: { width: 612, height: 792 }, LEGAL: { width: 612, height: 1008 }, TABLOID: { width: 792, height: 1224 }
};
export interface PdfSettings {
  readonly options: PdfOptions & Pick<PdfDocumentOptions, "paragraphStyles">;
  readonly paragraphStyles: Readonly<Record<string, Readonly<PdfParagraphStyle>>>;
  readonly page: PdfPageSize;
  readonly margins: PdfMargins;
  readonly fontSize: number;
  readonly padding: number;
  readonly limits: Required<Pick<PdfLimits, "maxPages" | "maxColumns" | "maxCellCharacters" | "maxTextRuns" | "maxGroupBlocks" | "maxRowLines" | "maxFontBytes" | "maxPageBytes" | "maxHyperlinks" | "maxImageBytes" | "maxImagePixels" | "maxVectorCommands" | "maxDrawingElements" | "maxStructureElements" | "maxMarkedContentReferences" | "maxListItems" | "maxParagraphStyles" | "maxSections" | "maxBookmarks">> & Pick<PdfLimits, "maxImageWorkingBytes">;
  readonly budget: ExportBudget;
}
export function settings(input: PdfOptions & Pick<PdfDocumentOptions, "paragraphStyles">, document = false): PdfSettings {
  const configuration = { ...input };
  if (configuration.limits) configuration.limits = Object.freeze({ ...configuration.limits });
  if (configuration.pageSize && typeof configuration.pageSize === "object") configuration.pageSize = Object.freeze({ ...configuration.pageSize });
  if (configuration.margins && typeof configuration.margins === "object") configuration.margins = Object.freeze({ ...configuration.margins });
  const textShaper = configuration.textShaper, shapeLine = textShaper?.shapeLine;
  if (textShaper !== undefined && (!textShaper || typeof shapeLine !== "function")) throw new TypeError("textShaper must provide shapeLine.");
  if (configuration.textDirection !== undefined && !["auto", "ltr", "rtl"].includes(configuration.textDirection))
    throw new TypeError("textDirection must be auto, ltr or rtl.");
  const columns = copyColumns(configuration?.columns).map(c => Object.freeze(c));
  if (!columns.length && !document) throw new TypeError("PDF tables need at least one declared column.");
  function copySpans(rows: TableSpanRows): TableSpanRows {
    tableSpans(rows, columns.length);
    return Object.freeze(rows.map(row => Object.freeze(row.map(cell => cell ? Object.freeze({ ...cell }) : null))));
  }
  const options: PdfOptions & Pick<PdfDocumentOptions, "paragraphStyles"> = { ...configuration, columns,
    ...(textShaper ? { textShaper: Object.freeze({ shapeLine: shapeLine!.bind(textShaper) }) } : {}),
    ...(configuration.tagged ? { tagged: Object.freeze({ ...configuration.tagged }) } : {}),
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
  const limits = { maxPages: 10000, maxColumns: 1024, maxCellCharacters: 1000000, maxTextRuns: 10000, maxGroupBlocks: 256, maxRowLines: 100000, maxFontBytes: 16 * 1024 * 1024, maxPageBytes: 8 * 1024 * 1024, maxHyperlinks: 100000,
    maxImageBytes: 32 * 1024 * 1024, maxImagePixels: 16 * 1024 * 1024, maxVectorCommands: 100000, maxDrawingElements: 100000,
    maxStructureElements: 1000000, maxMarkedContentReferences: 2000000,
    maxListItems: 100000, maxParagraphStyles: 256, maxSections: 1000, maxBookmarks: 10000 };
  for (const key of Object.keys(limits) as (keyof typeof limits)[]) {
    const override = options.limits?.[key]; if (override !== undefined) limits[key] = override;
  }
  const paragraphStyles = captureParagraphStyles(options.paragraphStyles, limits.maxParagraphStyles);
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
  for (const key of ["formatValue", "formatRuns"] as const)
    if (options[key] !== undefined && typeof options[key] !== "function") throw new TypeError(key + " must be a synchronous function.");
  for (const key of ["compression", "includeHeader", "pageNumbers"] as const) if (options[key] !== undefined && typeof options[key] !== "boolean") throw new TypeError(key + " must be boolean.");
  if (options.alternateRowColor !== undefined) color(options.alternateRowColor);
  if (options.footer?.values && options.footer.values.length > columns.length) throw new RangeError("Footer has more values than declared columns.");
  if (options.footer?.rows && (options.footer.values !== undefined || options.footer.totals !== undefined)) throw new TypeError("Structured footer rows cannot be combined with values/totals.");
  if (options.headerRows && options.includeHeader === false) throw new TypeError("headerRows requires includeHeader.");
  if (options.tagged !== undefined) {
    if (!options.tagged || typeof options.tagged !== "object" || typeof options.tagged.language !== "string" || options.tagged.language.length > 63)
      throw new TypeError("tagged.language must be a language tag such as en or en-US.");
    try { if (Intl.getCanonicalLocales(options.tagged.language).length !== 1) throw new RangeError(); }
    catch { throw new TypeError("tagged.language must be a language tag such as en or en-US."); }
    if (options.headerRows || columns.some(column => column.groups?.length) || options.footer?.rows)
      throw new TypeError("Tagged PDF supports leaf column headings and ordinary footer cells; grouped headings and span matrices are unsupported.");
  }
  const maxImageWorkingBytes = options.limits?.maxImageWorkingBytes;
  return { options: { ...options, paragraphStyles }, paragraphStyles, page, margins, fontSize, padding,
    limits: { ...limits, ...(maxImageWorkingBytes === undefined ? {} : { maxImageWorkingBytes }) }, budget };
}

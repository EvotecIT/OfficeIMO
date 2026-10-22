import { OfficeIMOError } from "../core/errors.js";
import { ExportBudget } from "../core/limits.js";
import { copyColumns } from "../internal/rows.js";
import { captureTableSpans } from "../internal/table-spans.js";
import { copyExportValue } from "../internal/export-values.js";
import { captureFields } from "../internal/capture.js";
import type { PdfOptions, PdfMargins, PdfPageSize, PdfLimits, PdfTableFooter } from "./types.js";
import type { PdfDocumentOptions, PdfParagraphStyle } from "./document-types.js";
import { captureParagraphStyles } from "./paragraph-styles.js";
import { captureBooleanText } from "../core/value-text.js";
import { positive, color, presentation } from "./presentation-settings.js";
export { positive, color, presentation } from "./presentation-settings.js";
const sizes: Readonly<Record<string, PdfPageSize>> = {
  A3: { width: 841.89, height: 1190.551 }, A4: { width: 595.276, height: 841.89 }, A5: { width: 419.528, height: 595.276 },
  LETTER: { width: 612, height: 792 }, LEGAL: { width: 612, height: 1008 }, TABLOID: { width: 792, height: 1224 }
};
/** @internal Shared capture schema for thin format adapters and the grouped writer. */
export const pdfOptionFields = ["columns", "signal", "onProgress", "limits", "textShaper", "textDirection", "pageSize", "orientation", "margins", "fonts", "tagged", "fontSize",
  "title", "messageTop", "messageBottom", "includeHeader", "headerRows", "footer", "headerPresentation", "footerPresentation", "alternateRowColor", "padding",
  "columnWidths", "wideTable", "pageHeader", "pageFooter", "pageNumbers", "pageNumberLabels", "booleanText", "formatValue", "formatRuns", "compression", "paragraphStyles"] as const;
/** @internal Complete PDF limit schema used when a host combines format and shared limits. */
export const pdfLimitFields = ["maxRows", "maxCells", "maxTextCharacters", "maxOutputBytes", "maxPages", "maxColumns", "maxTableBands", "maxTableWindowRows", "maxCellCharacters", "maxTextRuns", "maxGroupBlocks", "maxRowLines",
  "maxFontBytes", "maxPageBytes", "maxHyperlinks", "maxImageBytes", "maxImagePixels", "maxImageWorkingBytes", "maxVectorCommands", "maxDrawingElements", "maxStructureElements",
  "maxMarkedContentReferences", "maxSections", "maxBookmarks", "maxListItems", "maxParagraphStyles"] as const;

/** @internal Capture supported structural properties once, including class-based options with prototype accessors. */
export function capturePdfOptions<T extends object>(input: T): T {
  if (!input || typeof input !== "object" || Array.isArray(input)) throw new TypeError("PDF options must be an object.");
  const captured = captureFields(input, pdfOptionFields as unknown as readonly (keyof T)[]);
  const policy = (captured as unknown as PdfOptions).wideTable;
  return { ...captured, ...(policy === undefined ? {} : { wideTable: captureWideTable(policy) }) };
}
/** @internal Capture the complete PDF budget before an adapter adds its shared row/output limits. */
export function capturePdfLimits(input: PdfLimits): Readonly<PdfLimits> {
  if (!input || typeof input !== "object" || Array.isArray(input)) throw new TypeError("PDF limits must be an object.");
  const captured = captureFields(input, pdfLimitFields);
  for (const [key, value] of Object.entries(captured))
    if (value !== undefined && (!Number.isSafeInteger(value) || value < 0)) throw new RangeError(key + " must be a nonnegative safe integer.");
  return Object.freeze(captured);
}
function captureFooter(input: PdfTableFooter, columns: number): Readonly<PdfTableFooter> {
  if (!input || typeof input !== "object" || Array.isArray(input)) throw new TypeError("PDF footer must be an object.");
  const { values, totals, rows } = input;
  if (values !== undefined && !Array.isArray(values)) throw new TypeError("PDF footer values must be an array.");
  return Object.freeze({ ...(values === undefined ? {} : { values: Object.freeze(values.map(copyExportValue)) }),
    ...(totals === undefined ? {} : { totals: Object.freeze({ ...totals }) }),
    ...(rows === undefined ? {} : { rows: captureTableSpans(rows, columns) }) });
}
/** @internal Capture a band policy and its array before asynchronous layout or row production. */
export function captureWideTable(input: PdfOptions["wideTable"]): PdfOptions["wideTable"] {
  if (input === undefined || input === "fit" || input === "reject") return input;
  if (!input || typeof input !== "object" || Array.isArray(input) || Object.keys(input).some(key => !["mode", "minimumColumnWidth", "repeatColumns"].includes(key)))
    throw new TypeError("wideTable must be fit, reject or a bands policy.");
  const { mode, minimumColumnWidth, repeatColumns } = input;
  if (mode !== "bands") throw new TypeError("wideTable.mode must be bands.");
  if (minimumColumnWidth !== undefined) positive(minimumColumnWidth, "Minimum PDF column width");
  if (repeatColumns !== undefined && !Array.isArray(repeatColumns)) throw new TypeError("repeatColumns must be an array.");
  const repeats = repeatColumns === undefined ? [] : [...repeatColumns];
  if (repeats.some((value, index) => !Number.isSafeInteger(value) || value < 0 || index > 0 && value <= repeats[index - 1]!))
    throw new RangeError("repeatColumns must contain sorted unique nonnegative indices.");
  return Object.freeze({ mode, ...(minimumColumnWidth === undefined ? {} : { minimumColumnWidth }), repeatColumns: Object.freeze(repeats) });
}
export interface PdfSettings {
  readonly options: PdfOptions & Pick<PdfDocumentOptions, "paragraphStyles">;
  readonly paragraphStyles: Readonly<Record<string, Readonly<PdfParagraphStyle>>>;
  readonly page: PdfPageSize;
  readonly margins: PdfMargins;
  readonly fontSize: number;
  readonly padding: number;
  readonly limits: Required<Pick<PdfLimits, "maxPages" | "maxColumns" | "maxTableBands" | "maxTableWindowRows" | "maxCellCharacters" | "maxTextRuns" | "maxGroupBlocks" | "maxRowLines" | "maxFontBytes" | "maxPageBytes" | "maxHyperlinks" | "maxImageBytes" | "maxImagePixels" | "maxVectorCommands" | "maxDrawingElements" | "maxStructureElements" | "maxMarkedContentReferences" | "maxListItems" | "maxParagraphStyles" | "maxSections" | "maxBookmarks">> & Pick<PdfLimits, "maxImageWorkingBytes">;
  readonly budget: ExportBudget;
}
export function settings(input: PdfOptions & Pick<PdfDocumentOptions, "paragraphStyles">, document = false): PdfSettings {
  const configuration = { ...capturePdfOptions(input) };
  configuration.booleanText = captureBooleanText(configuration.booleanText);
  if (configuration.limits !== undefined) configuration.limits = capturePdfLimits(configuration.limits);
  if (configuration.pageSize && typeof configuration.pageSize === "object") configuration.pageSize = Object.freeze(captureFields(configuration.pageSize, ["width", "height"]));
  if (configuration.margins && typeof configuration.margins === "object") configuration.margins = Object.freeze(captureFields(configuration.margins, ["top", "right", "bottom", "left"]));
  if (configuration.pageNumberLabels !== undefined) {
    const labels = configuration.pageNumberLabels;
    if (!labels || typeof labels !== "object" || Array.isArray(labels)) throw new TypeError("pageNumberLabels must provide literal text.");
    const { prefix = "Page ", separator = " of ", suffix = "" } = labels;
    if ([prefix, separator, suffix].some(value => typeof value !== "string")) throw new TypeError("Page-number labels must be text.");
    configuration.pageNumberLabels = Object.freeze({ prefix, separator, suffix });
  }
  const textShaper = configuration.textShaper, shapeLine = textShaper?.shapeLine;
  if (textShaper !== undefined && (!textShaper || typeof shapeLine !== "function")) throw new TypeError("textShaper must provide shapeLine.");
  if (configuration.textDirection !== undefined && !["auto", "ltr", "rtl"].includes(configuration.textDirection))
    throw new TypeError("textDirection must be auto, ltr or rtl.");
  const columns = document ? [] : copyColumns(configuration.columns).map(c => Object.freeze(c));
  if (document) configuration.includeHeader = false;
  if (!columns.length && !document) throw new TypeError("PDF tables need at least one declared column.");
  const capturedFonts = configuration.fonts === undefined ? undefined : captureFields(configuration.fonts, ["regular", "bold", "italic", "boldItalic"]);
  const options: PdfOptions & Pick<PdfDocumentOptions, "paragraphStyles"> = { ...configuration, columns,
    ...(textShaper ? { textShaper: Object.freeze({ shapeLine: shapeLine!.bind(textShaper) }) } : {}),
    ...(configuration.tagged ? { tagged: Object.freeze(captureFields(configuration.tagged, ["language"])) } : {}),
    ...(configuration.headerRows ? { headerRows: captureTableSpans(configuration.headerRows, columns.length) } : {}),
    ...(capturedFonts ? { fonts: Object.freeze(capturedFonts) } : {}),
    ...(configuration.columnWidths ? { columnWidths: Object.freeze([...configuration.columnWidths]) } : {}),
    ...(configuration.footer === undefined ? {} : { footer: captureFooter(configuration.footer, columns.length) }),
    headerPresentation: Object.freeze({ background: "e7edf5", bold: true, ...presentation(configuration.headerPresentation) }),
    footerPresentation: Object.freeze({ background: "eef2f6", bold: true, ...presentation(configuration.footerPresentation) }) };
  const budget = new ExportBudget(options.limits);
  const limits = { maxPages: 10000, maxColumns: 1024, maxTableBands: 64, maxTableWindowRows: 256, maxCellCharacters: 1000000, maxTextRuns: 10000, maxGroupBlocks: 256, maxRowLines: 100000, maxFontBytes: 16 * 1024 * 1024, maxPageBytes: 8 * 1024 * 1024, maxHyperlinks: 100000,
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
  }
  const maxImageWorkingBytes = options.limits?.maxImageWorkingBytes;
  return { options: { ...options, paragraphStyles }, paragraphStyles, page, margins, fontSize, padding,
    limits: { ...limits, ...(maxImageWorkingBytes === undefined ? {} : { maxImageWorkingBytes }) }, budget };
}

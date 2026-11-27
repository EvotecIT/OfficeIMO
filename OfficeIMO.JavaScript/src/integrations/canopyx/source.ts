import { ExportCell, checkAbort, inputRows } from "../../core/index.js";
import type { CellPresentation, CellValue, ExportValue } from "../../core/index.js";
import { canopyPivot, pivotNumberFormat } from "./pivot.js";
import type { CanopyPivotMapping } from "./pivot.js";
import { datetime } from "./dates.js";
import { captureFields } from "../../internal/capture.js";
import { copyColumns } from "../../internal/rows.js";
import { copyCellPresentation } from "../../core/presentation.js";
import { copyExportLink } from "../../core/links.js";
import { cleanXml } from "../../xml/index.js";
import { canopyDiagnostics } from "./diagnostics.js";
import { canopyNumberFormat, canopyXlsxWidth } from "./presentation.js";
import { canopyDefaultTones } from "./tones.js";
import type { XlsxOptions } from "../../xlsx/index.js";
import type { CanopyCapture, CanopyDiagnostic, CanopyExport, CanopyFormat, CanopySourceOptions } from "./types.js";

const columnFields = ["width", "format", "wrapText", "alignment", "rowHeader", "groups"] as const;
// Captured presentations whose cell text already carries the readable content (badge labels are appended below).
// Their visual loss is reported, never treated as an unsupported export.
const textFallbacks: ReadonlySet<string> = new Set(["relative-link", "history-series", "badges", "custom-class", "custom-renderer"]);
const badgeSeparator = " \u00b7 ";
const rawIdentities = new WeakMap<ExportCell, CellValue>();
// Appended badge labels, kept apart so later passes that re-render a cell's text (PDF numbers) can carry them over.
const badgeSuffixes = new WeakMap<ExportCell, string>();

/** @internal The text a later pass may re-render, and the badge suffix it must append back unchanged. */
export function canopyBadgeSplit(cell: ExportCell): { readonly text: string; readonly suffix: string } {
  const suffix = badgeSuffixes.get(cell), text = cell.text ?? "";
  return suffix !== undefined && text.endsWith(suffix) ? { text: text.slice(0, text.length - suffix.length), suffix } : { text, suffix: "" };
}

/** @internal Grouping uses captured identity even when a detail cell displays a relative-link label. */
export function canopyIdentityValue(value: ExportValue): CellValue {
  if (!(value instanceof ExportCell)) return value;
  const original = rawIdentities.get(value);
  return original === undefined ? value.value : original;
}

// Grids accept any number of badges. Exports keep the first labels that stay readable and fit an Excel text cell,
// and report how many were left out instead of failing the whole document.
const maximumBadges = 64, maximumCellText = 32767;
// Excel allows this many hyperlinks on one worksheet; later link cells keep their label without the link.
const maximumWorksheetLinks = 65530;

function badgeText(input: unknown, text: string): { readonly labels: string | undefined; readonly dropped: number } {
  if (!Array.isArray(input)) throw new TypeError("Canopy cell badges must be an array of { text, tone?, title? }.");
  let labels = "", kept = 0, dropped = 0;
  for (const badge of input) {
    const value = typeof badge === "string" ? badge : badge && typeof badge === "object" ? (badge as { text?: unknown }).text : undefined;
    if (typeof value !== "string") throw new TypeError("Canopy cell badges must be an array of { text, tone?, title? }.");
    const label = value.trim();
    if (!label) continue;
    const next = labels ? labels + badgeSeparator + label : label;
    if (kept >= maximumBadges || (text ? text.length + badgeSeparator.length : 0) + next.length > maximumCellText) { dropped++; continue; }
    labels = next; kept++;
  }
  return { labels: labels || undefined, dropped };
}

/** Map an immutable CanopyX record capture to a writer without querying the grid or retaining its body. */
export function createCanopyExport(capture: CanopyCapture, format: CanopyFormat, options: CanopySourceOptions = {}): CanopyExport {
  return prepareCanopyExport(capture, format, options).source;
}

/** @internal Capture the input once before full-writer metadata mapping and asynchronous enumeration. */
export function prepareCanopyExport(capture: CanopyCapture, format: CanopyFormat, options: CanopySourceOptions = {}): {
  readonly source: CanopyExport;
  readonly request: CanopyCapture["request"];
  readonly report: (diagnostic: CanopyDiagnostic) => void;
  readonly pivot: CanopyPivotMapping | undefined;
  readonly tones: ReadonlyMap<string, Readonly<CellPresentation>>;
  /** Text columns whose raw scalars must not carry badge labels, because an emitted native rule reads them. */
  readonly ruleOperands: Set<string>;
} {
  options = captureFields(options, ["tones", "columnOptions", "unsupportedPresentation", "datetime", "xlsx", "signal", "onDiagnostic"]);
  if (format !== "xlsx" && format !== "csv" && format !== "pdf") throw new TypeError("Canopy export format must be xlsx, csv or pdf.");
  if (!capture || typeof capture.rows !== "function" || !capture.request) throw new TypeError("A portable CanopyX capture with request and rows() is required.");
  const request = captureFields(capture.request, ["columns", "recordCount", "values", "presentation", "timeZone", "revision", "scope", "total", "matched", "selection", "query", "generatedAt", "grouping", "pivot", "formatRules"]);
  const pivot = canopyPivot(request);
  const signal = options.signal, values = request.values;
  if (!Array.isArray(request.columns) || !request.columns.length || (request.recordCount !== null && (!Number.isSafeInteger(request.recordCount) || request.recordCount < 0)) ||
    (values !== "raw" && values !== "display") || (request.timeZone !== "utc" && request.timeZone !== "local") || typeof request.revision !== "string")
    throw new TypeError("Invalid portable CanopyX capture metadata.");
  const dateMode = options.datetime ?? "preserve", report = canopyDiagnostics(format, options);
  // These captured presentations have a complete scalar/text projection. Report
  // their visual loss without treating an ordinary report export as unsupported.
  const fallback = canopyDiagnostics(format, { ...options, unsupportedPresentation: "text" });
  const clock = options.xlsx?.dateMode ?? request.timeZone;
  if (clock !== "utc" && clock !== "local") throw new TypeError("XLSX dateMode must be utc or local.");
  if (dateMode !== "preserve" && dateMode !== "typed" && dateMode !== "text") throw new TypeError("datetime must be preserve, typed or text.");
  const diagnostics = (items: readonly string[] | undefined, rowId?: string, columnId?: string): void => {
    if (items === undefined) return;
    if (!Array.isArray(items) || items.some(item => typeof item !== "string")) throw new TypeError("Canopy presentation diagnostics must be strings.");
    for (const code of items) {
      const diagnostic = { code, ...(rowId === undefined ? {} : { rowId }), ...(columnId === undefined ? {} : { columnId }) };
      (textFallbacks.has(code) ? fallback : report)(diagnostic);
    }
  };
  if (request.presentation !== "text" && request.presentation !== "semantic") throw new TypeError("Canopy captures must declare text or semantic presentation.");
  const tones = new Map(Object.entries(options.tones ?? {}).map(([name, style]) => {
    if (!style || typeof style !== "object" || Array.isArray(style)) throw new TypeError("Tone mappings must be portable presentation objects.");
    return [name, copyCellPresentation(style)] as const;
  }));
  const tone = (name: string | undefined, rowId: string, columnId?: string): Readonly<CellPresentation> | undefined => {
    if (name === undefined) return undefined;
    if (typeof name !== "string") throw new TypeError("Canopy tones must be strings.");
    const style = tones.get(name) ?? canopyDefaultTones.get(name);
    if (!style && name !== "neutral") fallback({ code: "UNMAPPED_TONE:" + name, rowId, ...(columnId === undefined ? {} : { columnId }) });
    return style;
  };
  // The full writer fills this with columns read by the native rules it emits, before any row is enumerated.
  const ruleOperands = new Set<string>();
  const ids = new Set<string>();
  const specs = (pivot?.columns ?? request.columns).map(input => {
    const fields = captureFields(input, ["id", "title", "kind", "width", "pinned", "format", "wrap", "alignment", "diagnostics"]);
    const column = Object.freeze({ ...fields,
      ...(fields.width === undefined ? {} : { width: Object.freeze(captureFields(fields.width, ["unit", "preferred", "minimum", "maximum"])) }),
      ...(fields.format === undefined ? {} : { format: Object.freeze(captureFields(fields.format, ["minimumFractionDigits", "maximumFractionDigits", "prefix", "suffix"])) }) });
    if (!column || typeof column.id !== "string" || !column.id || ids.has(column.id) || typeof column.title !== "string" ||
      !["text", "number", "datetime"].includes(column.kind) ||
      (column.wrap !== undefined && typeof column.wrap !== "boolean") || (column.alignment !== undefined && !["left", "right"].includes(column.alignment)) ||
      (column.width !== undefined && (column.width.unit !== "css-px" || !Number.isFinite(column.width.preferred) ||
        !Number.isFinite(column.width.minimum) || column.width.minimum <= 0 || column.width.preferred < column.width.minimum ||
        column.width.maximum !== undefined && (!Number.isFinite(column.width.maximum) || column.width.maximum < column.width.preferred))) ||
      (!pivot && request.presentation === "semantic" && (column.wrap === undefined || column.alignment === undefined || column.width === undefined)))
      throw new TypeError("Invalid or duplicate Canopy export column.");
    ids.add(column.id); diagnostics(column.diagnostics, undefined, column.id);
    const selected: NonNullable<CanopySourceOptions["columnOptions"]>[string] = Object.hasOwn(options.columnOptions ?? {}, column.id) ? options.columnOptions![column.id]! : {};
    for (const key of Object.keys(selected)) if (!(columnFields as readonly string[]).includes(key)) throw new TypeError("Unsupported Canopy column option: " + key);
    const override = captureFields(selected, columnFields);
    const numericFormat = (format === "xlsx" || format === "pdf") && values === "raw" && column.kind === "number" && column.format !== undefined && override.format === undefined
      ? canopyNumberFormat(column.format, report, column.id) : undefined;
    const [captured] = copyColumns([{ ...override, key: column.id, header: column.title,
      ...(format === "xlsx" && override.width === undefined && column.width !== undefined ? { width: canopyXlsxWidth(column.width.preferred) } : {}),
      ...(override.format === undefined && numericFormat !== undefined ? { format: numericFormat } : {}),
      ...(format === "xlsx" && values === "raw" && column.kind === "datetime" && override.format === undefined ? { format: "yyyy-mm-dd hh:mm:ss.000" } : {}),
      ...(values === "raw" && column.kind === "number" ? { type: "number" as const } : {}),
      // Screen nowrap is not a promise that its CSS width will survive PDF fitting.
      // PDF defaults to lossless wrapping; an explicit writer override stays strict.
      ...(format === "pdf" ? { wrapText: override.wrapText ?? true } :
        override.wrapText === undefined && column.wrap === undefined ? {} : { wrapText: override.wrapText ?? column.wrap! }),
      ...(override.alignment === undefined && column.alignment === undefined ?
        format === "pdf" && column.kind === "number" ? { alignment: "right" as const } : {} : { alignment: override.alignment ?? column.alignment! }) }]);
    return Object.freeze({ id: column.id, kind: column.kind, input: column, nativeFormatOverride: override.format, column: Object.freeze(captured!) });
  });
  const rowCount = request.recordCount;
  // The XLSX worksheet counts every hyperlink it holds: configured sheet links, the internal link of each preserved
  // oversized text cell, and captured cell links. Only captured cell links can fall back to their label.
  // writeCanopyTo passes its complete captured XLSX options through the narrower public source type.
  const xlsx = options.xlsx as Pick<XlsxOptions, "dateMode" | "oversizedText" | "invalidCharacterPolicy" | "sheet"> | undefined;
  const preserveText = format === "xlsx" && xlsx?.oversizedText === "preserve";
  const preserved = (value: unknown): boolean => preserveText && typeof value === "string" && value.length > maximumCellText &&
    cleanXml(value, xlsx?.invalidCharacterPolicy).length > maximumCellText;
  const sheetHyperlinks = xlsx?.sheet?.hyperlinks;
  const reservedLinks = (Array.isArray(sheetHyperlinks) ? sheetHyperlinks.length : 0) +
    (xlsx?.sheet?.includeHeader === false ? 0 : specs.filter(spec => preserved(spec.column.header)).length);
  const source: CanopyExport = Object.freeze({ columns: Object.freeze(specs.map(spec => spec.column)), rowCount, async *rows() {
    checkAbort(signal); let count = 0, links = reservedLinks, droppedLinks = 0;
    for await (const row of inputRows(capture.rows(signal ? { signal } : {}), signal)) {
      checkAbort(signal);
      if (!row || typeof row.id !== "string" || !row.cells || typeof row.cells !== "object") throw new TypeError("Invalid Canopy export record.");
      pivot?.observe(row);
      diagnostics(row.diagnostics, row.id);
      const rowStyle = tone(row.tone, row.id), result: ExportValue[] = [];
      for (const spec of specs) {
        if (!Object.hasOwn(row.cells, spec.id)) throw new TypeError("Canopy record is missing column " + spec.id + ".");
        const cell = row.cells[spec.id];
        if (!cell || typeof cell !== "object" || typeof cell.text !== "string" ||
          (cell.value !== null && !["string", "number", "boolean"].includes(typeof cell.value)) ||
          (typeof cell.value === "number" && (!Number.isFinite(cell.value) || Math.abs(cell.value) > Number.MAX_SAFE_INTEGER)))
          throw new TypeError("A portable Canopy GridExportCell with scalar value and resolved text is required.");
        diagnostics(cell.diagnostics, row.id, spec.id);
        // A relative destination is navigation metadata, not the visible label.
        // History captures already carry their frozen latest scalar (or null).
        // Columns read by an emitted native rule keep the raw scalar the rule compares, so the label goes to the text only.
        const relativeLabel = spec.kind === "text" && cell.diagnostics?.includes("relative-link") && !ruleOperands.has(spec.id);
        const badged = cell.badges === undefined ? undefined : badgeText(cell.badges, cell.text), badges = badged?.labels;
        if (badged?.dropped) fallback({ code: "BADGES_TRUNCATED:" + badged.dropped, rowId: row.id, columnId: spec.id });
        const text = badges ? cell.text ? cell.text + badgeSeparator + badges : badges : cell.text;
        // Badges belong to the visible label, so a badged text cell exports that composed label rather than a raw/label hybrid.
        // Typed columns keep their scalar, and native rule operands keep the raw value the rule compares; both carry badges in the text only.
        const badgedValue = !!badges && values === "raw" && !relativeLabel && spec.kind === "text" && !ruleOperands.has(spec.id);
        let value: import("../../core/index.js").CellValue = values === "display" || relativeLabel || badgedValue ? text : cell.value;
        if (format === "xlsx" && values === "raw" && spec.kind === "datetime" && typeof value === "string") value = datetime(value, dateMode, clock);
        const capturedFormat = pivot && spec.kind === "number" && cell.value !== null ? pivotNumberFormat(cell) : undefined;
        const nativeFormat = spec.nativeFormatOverride === "" ? "General" : spec.nativeFormatOverride ?? capturedFormat;
        const summary = format === "pdf" && pivot && (row.summary || row.grandTotal);
        const cellStyle = tone(cell.tone, row.id, spec.id), presentation = rowStyle || cellStyle || nativeFormat || summary ? {
          ...rowStyle, ...(summary ? { bold: true } : {}), ...cellStyle,
          ...((format === "pdf" || (format === "xlsx" && pivot)) && values === "raw" && nativeFormat !== undefined ? { numberFormat: nativeFormat } : {}) } : undefined;
        // CSV has no presentation or link metadata: emit its selected scalar directly.
        // This also avoids a frozen wrapper for every cell in large CSV captures.
        if (format === "csv") result.push(value);
        else {
          let link = cell.link;
          if (format === "xlsx") {
            if (preserved(value)) links++;
            // ExportCell validates the links it keeps; a link left out is still validated, so an unsafe target is never hidden.
            if (link !== undefined && ++links > maximumWorksheetLinks) { copyExportLink(link); link = undefined; links--; droppedLinks++; }
          }
          const exported = new ExportCell(value, { text, ...(presentation ? { presentation } : {}), ...(link === undefined ? {} : { link }) });
          if (relativeLabel || badgedValue) rawIdentities.set(exported, cell.value);
          // A badge-only label is all suffix: later passes then see the captured label as empty.
          if (badges) badgeSuffixes.set(exported, cell.text ? badgeSeparator + badges : badges);
          result.push(exported);
        }
      }
      if (rowCount !== null && count >= rowCount) throw new TypeError("Canopy capture exceeded its declared recordCount.");
      count++; yield result;
    }
    if (rowCount !== null && count !== rowCount) throw new TypeError("Canopy capture did not produce its declared recordCount.");
    if (droppedLinks) fallback({ code: "LINKS_TRUNCATED:" + droppedLinks });
  } });
  return Object.freeze({ source, request: Object.freeze({ ...request, columns: Object.freeze(specs.map(spec => spec.input)) }), report, tones, pivot, ruleOperands });
}

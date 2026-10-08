import { ExportCell, checkAbort, inputRows } from "../../core/index.js";
import type { CellPresentation, ExportValue } from "../../core/index.js";
import { datetime } from "./dates.js";
import type { CanopyCapture, CanopyDiagnostic, CanopyExport, CanopyFormat, CanopySourceOptions } from "./types.js";

/** Map an immutable CanopyX record capture to a writer without querying the grid or retaining its body. */
export function createCanopyExport(capture: CanopyCapture, format: CanopyFormat, options: CanopySourceOptions = {}): CanopyExport {
  if (format !== "xlsx" && format !== "csv" && format !== "pdf") throw new TypeError("Canopy export format must be xlsx, csv or pdf.");
  if (!capture || typeof capture.rows !== "function" || !capture.request) throw new TypeError("A portable CanopyX capture with request and rows() is required.");
  const { request } = capture, signal = options.signal;
  if (!Array.isArray(request.columns) || !request.columns.length || (request.recordCount !== null && (!Number.isSafeInteger(request.recordCount) || request.recordCount < 0)) ||
    (request.values !== "raw" && request.values !== "display") || (request.timeZone !== "utc" && request.timeZone !== "local") || typeof request.revision !== "string")
    throw new TypeError("Invalid portable CanopyX capture metadata.");
  const policy = options.unsupportedPresentation ?? (format === "csv" ? "text" : "reject"), dateMode = options.datetime ?? "preserve", notify = options.onDiagnostic;
  const clock = options.xlsx?.dateMode ?? request.timeZone;
  if (clock !== "utc" && clock !== "local") throw new TypeError("XLSX dateMode must be utc or local.");
  if (policy !== "reject" && policy !== "text") throw new TypeError("unsupportedPresentation must be reject or text.");
  if (dateMode !== "preserve" && dateMode !== "typed" && dateMode !== "text") throw new TypeError("datetime must be preserve, typed or text.");
  if (notify !== undefined && typeof notify !== "function") throw new TypeError("onDiagnostic must be a synchronous function.");
  const report = (diagnostic: CanopyDiagnostic): void => {
    if (policy === "reject") throw new TypeError("Unsupported Canopy presentation: " + diagnostic.code + (diagnostic.columnId ? " (" + diagnostic.columnId + ")" : ""));
    const result: unknown = notify?.(Object.freeze(diagnostic));
    if (result && typeof (result as { then?: unknown }).then === "function") { void Promise.resolve(result).catch(() => {}); throw new TypeError("onDiagnostic must be synchronous."); }
  };
  const diagnostics = (items: readonly string[] | undefined, rowId?: string, columnId?: string): void => {
    if (items === undefined) return;
    if (!Array.isArray(items) || items.some(item => typeof item !== "string")) throw new TypeError("Canopy presentation diagnostics must be strings.");
    for (const code of items) report({ code, ...(rowId === undefined ? {} : { rowId }), ...(columnId === undefined ? {} : { columnId }) });
  };
  if (request.presentation === "text") report({ code: "TEXT_ONLY_CAPTURE" });
  else if (request.presentation !== "semantic") throw new TypeError("Canopy captures must declare text or semantic presentation.");
  const tones = new Map(Object.entries(options.tones ?? {}).map(([name, style]) => {
    if (!style || typeof style !== "object" || Array.isArray(style)) throw new TypeError("Tone mappings must be portable presentation objects.");
    return [name, Object.freeze({ ...style })] as const;
  }));
  const tone = (name: string | undefined, rowId: string, columnId?: string): Readonly<CellPresentation> | undefined => {
    if (name === undefined) return undefined;
    if (typeof name !== "string") throw new TypeError("Canopy tones must be strings.");
    const style = tones.get(name);
    if (!style) report({ code: "UNMAPPED_TONE:" + name, rowId, ...(columnId === undefined ? {} : { columnId }) });
    return style;
  };
  const ids = new Set<string>();
  const specs = request.columns.map(column => {
    if (!column || typeof column.id !== "string" || !column.id || ids.has(column.id) || typeof column.title !== "string" ||
      !["text", "number", "datetime"].includes(column.kind) ||
      (column.wrap !== undefined && typeof column.wrap !== "boolean") || (column.alignment !== undefined && !["left", "right"].includes(column.alignment)) ||
      (column.width !== undefined && (column.width.unit !== "css-px" || !Number.isFinite(column.width.preferred) ||
        !Number.isFinite(column.width.minimum) || column.width.minimum <= 0 || column.width.preferred < column.width.minimum ||
        column.width.maximum !== undefined && (!Number.isFinite(column.width.maximum) || column.width.maximum < column.width.preferred))) ||
      (request.presentation === "semantic" && (column.wrap === undefined || column.alignment === undefined || column.width === undefined)))
      throw new TypeError("Invalid or duplicate Canopy export column.");
    ids.add(column.id); diagnostics(column.diagnostics, undefined, column.id);
    const override = Object.hasOwn(options.columnOptions ?? {}, column.id) ? options.columnOptions![column.id]! : {};
    for (const key of Object.keys(override)) if (!["width", "format", "wrapText", "alignment", "groups"].includes(key)) throw new TypeError("Unsupported Canopy column option: " + key);
    return Object.freeze({ id: column.id, kind: column.kind, column: Object.freeze({ ...override, key: column.id, header: column.title,
      ...(override.wrapText === undefined && column.wrap === undefined ? {} : { wrapText: override.wrapText ?? column.wrap! }),
      ...(override.alignment === undefined && column.alignment === undefined ? {} : { alignment: override.alignment ?? column.alignment! }),
      ...(override.groups === undefined ? {} : { groups: Object.freeze([...override.groups]) }) }) });
  });
  const values = request.values, rowCount = request.recordCount;
  return Object.freeze({ columns: Object.freeze(specs.map(spec => spec.column)), rowCount, async *rows() {
    checkAbort(signal); let count = 0;
    for await (const row of inputRows(capture.rows(signal ? { signal } : {}), signal)) {
      checkAbort(signal);
      if (!row || typeof row.id !== "string" || !row.cells || typeof row.cells !== "object") throw new TypeError("Invalid Canopy export record.");
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
        let value: import("../../core/index.js").CellValue = values === "display" ? cell.text : cell.value;
        if (format === "xlsx" && values === "raw" && spec.kind === "datetime" && typeof value === "string") value = datetime(value, dateMode, clock);
        const cellStyle = tone(cell.tone, row.id, spec.id), presentation = rowStyle || cellStyle ? { ...rowStyle, ...cellStyle } : undefined;
        // CSV has no presentation or link metadata: emit its selected scalar directly.
        // This also avoids a frozen wrapper for every cell in large CSV captures.
        result.push(format === "csv" ? value : new ExportCell(value, { text: cell.text, ...(presentation ? { presentation } : {}), ...(cell.link === undefined ? {} : { link: cell.link }) }));
      }
      if (rowCount !== null && count >= rowCount) throw new TypeError("Canopy capture exceeded its declared recordCount.");
      count++; yield result;
    }
    if (rowCount !== null && count !== rowCount) throw new TypeError("Canopy capture did not produce its declared recordCount.");
  } });
}

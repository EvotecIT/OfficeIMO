import { ExportCell, NotSupportedError, OfficeIMOError } from "../core/index.js";
import { assertExportValue } from "../core/presentation.js";
import { isDate } from "../core/dates.js";
import type { CellValue, ErrorCode, ExportProgress, ExportResult, ExportValue } from "../core/index.js";
import type { ExportCellOptions } from "../core/presentation.js";
import type { WorkerExportOptions } from "./types.js";

export const protocol = "@evotecit/officeimo/worker-export/1";
export interface Bounds {
  readonly batchRows: number; readonly maxBatchCells: number; readonly maxBatchCharacters: number; readonly maxChunkBytes: number;
  readonly maxSetupBytes: number; readonly maxSetupNodes: number;
}
export interface WireCell { readonly value: CellValue; readonly options?: ExportCellOptions; }
export interface WireMessage {
  readonly protocol: typeof protocol;
  readonly job: string;
  readonly kind: string;
  readonly id?: number;
  readonly input?: unknown;
  readonly bounds?: Bounds;
  readonly rows?: readonly (readonly WireCell[])[];
  readonly done?: boolean;
  readonly bytes?: Uint8Array;
  readonly progress?: ExportProgress;
  readonly result?: ExportResult;
  readonly error?: { readonly name: string; readonly message: string; readonly code?: ErrorCode; readonly feature?: string };
}
export function message(value: unknown): WireMessage | undefined {
  if (!value || typeof value !== "object") return undefined;
  const candidate = value as WireMessage;
  return candidate.protocol === protocol && typeof candidate.job === "string" && typeof candidate.kind === "string" ? candidate : undefined;
}
export function bounds(options: WorkerExportOptions): Bounds {
  const result = { batchRows: options.batchRows ?? 64, maxBatchCells: options.maxBatchCells ?? 65536,
    maxBatchCharacters: options.maxBatchCharacters ?? 1000000, maxChunkBytes: options.maxChunkBytes ?? 65536,
    maxSetupBytes: options.maxSetupBytes ?? 16 * 1024 * 1024, maxSetupNodes: options.maxSetupNodes ?? 100000 };
  for (const [name, value] of Object.entries(result))
    if (!Number.isSafeInteger(value) || value < 1) throw new RangeError(name + " must be a positive safe integer.");
  return Object.freeze(result);
}
/** Snapshot portable setup before messaging. Font objects, functions, accessors and cyclic live objects stay in their host. */
export function captureInput(input: unknown, limits: Bounds): unknown {
  let nodes = 0, bytes = 0;
  const ancestors = new Set<object>();
  function accept(additional = 0): void {
    bytes += additional;
    if (bytes > limits.maxSetupBytes || ++nodes > limits.maxSetupNodes)
      throw new OfficeIMOError("RESOURCE_LIMIT", "Worker configuration exceeds its message budget.");
  }
  function copy(value: unknown, depth: number): unknown {
    if (depth > 32) throw new RangeError("Worker configuration nesting exceeds 32 levels.");
    accept(typeof value === "string" ? value.length * 2 : 0);
    if (value == null || ["string", "number", "boolean"].includes(typeof value)) return value;
    if (isDate(value)) { accept(8); return new Date(value.getTime()); }
    if (value instanceof Uint8Array) { accept(value.byteLength); return new Uint8Array(value); }
    if (typeof value !== "object") throw new TypeError("Worker configuration requires portable data without functions or symbols.");
    if (ancestors.has(value)) throw new TypeError("Worker configuration cannot contain cycles.");
    const prototype: unknown = Object.getPrototypeOf(value);
    if (!Array.isArray(value) && prototype !== Object.prototype && prototype !== null)
      throw new TypeError("Worker configuration supports plain records, arrays, Dates and Uint8Array bytes; reconstruct font instances in the worker.");
    ancestors.add(value);
    try {
      if (Array.isArray(value)) {
        if (value.length > limits.maxSetupNodes - nodes) throw new OfficeIMOError("RESOURCE_LIMIT", "Worker configuration exceeds its message budget.");
        const snapshot: unknown[] = [];
        for (let i = 0; i < value.length; i++) {
          const descriptor = Object.getOwnPropertyDescriptor(value, String(i));
          if (descriptor && !Object.prototype.hasOwnProperty.call(descriptor, "value")) throw new TypeError("Worker configuration cannot contain accessors.");
          snapshot.push(copy(descriptor?.value, depth + 1));
        }
        return snapshot;
      }
      const snapshot: Record<string, unknown> = {};
      for (const key of Reflect.ownKeys(value)) {
        if (typeof key !== "string") throw new TypeError("Worker configuration keys must be strings.");
        const descriptor = Object.getOwnPropertyDescriptor(value, key)!;
        if (!descriptor.enumerable) continue;
        if (!Object.prototype.hasOwnProperty.call(descriptor, "value")) throw new TypeError("Worker configuration cannot contain accessors.");
        accept(key.length * 2);
        Object.defineProperty(snapshot, key, { value: copy(descriptor.value, depth + 1), enumerable: true, configurable: true, writable: true });
      }
      return snapshot;
    } finally { ancestors.delete(value); }
  }
  return copy(input, 0);
}
export function captureRow(row: readonly ExportValue[]): { readonly cells: readonly WireCell[]; readonly characters: number } {
  if (!Array.isArray(row)) throw new TypeError("Worker export rows must be projected arrays.");
  let characters = 0;
  const scalar = (value: CellValue): CellValue => {
    if (typeof value === "string") characters += value.length;
    return isDate(value) ? new Date(value.getTime()) : value;
  };
  const cells = row.map(value => {
    assertExportValue(value);
    if (!(value instanceof ExportCell)) return { value: scalar(value) };
    const options: ExportCellOptions = {
      ...(value.text === undefined ? {} : { text: value.text }),
      ...(value.presentation === undefined ? {} : { presentation: { ...value.presentation } }),
      ...(value.link === undefined ? {} : { link: { ...value.link } })
    };
    characters += value.text?.length ?? 0;
    for (const entry of Object.values(options.presentation ?? {})) if (typeof entry === "string") characters += entry.length;
    characters += (value.link?.target.length ?? 0) + (value.link?.tooltip?.length ?? 0);
    return { value: scalar(value.value), options };
  });
  return { cells, characters };
}
export function restoreRows(rows: WireMessage["rows"], limits: Bounds): readonly (readonly ExportValue[])[] {
  if (!Array.isArray(rows) || rows.length > limits.batchRows) throw new TypeError("Invalid worker row batch.");
  let cells = 0, characters = 0;
  return rows.map(row => {
    if (!Array.isArray(row)) throw new TypeError("Invalid worker row.");
    if (cells + row.length > limits.maxBatchCells) throw new OfficeIMOError("RESOURCE_LIMIT", "Worker row batch exceeds its message budget.");
    const restored = row.map((cell: WireCell) => {
      if (!cell || typeof cell !== "object") throw new TypeError("Invalid worker cell.");
      const value = cell.options === undefined ? cell.value : new ExportCell(cell.value, cell.options);
      assertExportValue(value); return value;
    });
    const captured = captureRow(restored); cells += row.length; characters += captured.characters;
    if (cells > limits.maxBatchCells || characters > limits.maxBatchCharacters) throw new OfficeIMOError("RESOURCE_LIMIT", "Worker row batch exceeds its message budget.");
    return restored;
  });
}
export function captureProgress(value: ExportProgress, limits: Bounds): ExportProgress {
  if (!value || !["rows", "complete"].includes(value.phase) || !Number.isSafeInteger(value.rows) || value.rows < 0)
    throw new TypeError("Invalid worker progress.");
  for (const key of ["bytes", "sheetRows", "totalRows"] as const)
    if (value[key] !== undefined && (!Number.isSafeInteger(value[key]) || value[key]! < 0)) throw new TypeError("Invalid worker progress count.");
  if (value.sheetName !== undefined && (typeof value.sheetName !== "string" || value.sheetName.length > limits.maxBatchCharacters))
    throw new TypeError("Invalid or oversized worker sheet name.");
  return { phase: value.phase, rows: value.rows,
    ...(value.bytes === undefined ? {} : { bytes: value.bytes }), ...(value.sheetRows === undefined ? {} : { sheetRows: value.sheetRows }),
    ...(value.totalRows === undefined ? {} : { totalRows: value.totalRows }), ...(value.sheetName === undefined ? {} : { sheetName: value.sheetName }) };
}
export function wireError(error: unknown): NonNullable<WireMessage["error"]> {
  return { name: error instanceof Error ? error.name : "Error", message: error instanceof Error ? error.message : String(error),
    ...(error instanceof OfficeIMOError ? { code: error.code } : {}), ...(error instanceof NotSupportedError ? { feature: error.feature } : {}) };
}
export function restoreError(error: WireMessage["error"]): Error {
  if (!error || typeof error.message !== "string") return new Error("Worker export failed.");
  const result = error.code === "NOT_SUPPORTED" && typeof error.feature === "string" ? new NotSupportedError(error.feature)
    : error.code === undefined ? new Error(error.message) : new OfficeIMOError(error.code, error.message);
  result.message = error.message;
  if (typeof error.name === "string") result.name = error.name;
  return result;
}
export function result(value: ExportResult | undefined): ExportResult {
  if (!value || [value.rows, value.columns, value.bytes].some(n => !Number.isSafeInteger(n) || n < 0)) throw new TypeError("Invalid worker export result.");
  return { rows: value.rows, columns: value.columns, bytes: value.bytes };
}

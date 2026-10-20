import type { CellValue, Column, ExportLimits } from "./index.js";
import { ExportCell, assertExportValue } from "./presentation.js";
import type { ExportValue } from "./presentation.js";
import { isDate } from "./dates.js";
import { checkAbort, inputRows, pause, taskYieldDue } from "./iteration.js";
import { ExportBudget } from "./limits.js";
import { OfficeIMOError } from "./errors.js";
import { copyColumns, createRowProjector } from "../internal/rows.js";
import type { ProjectionColumn } from "../internal/columns.js";
import { createTotals } from "../internal/totals.js";
import type { NumericAggregate, TotalOperation } from "../internal/totals.js";

export type { TotalOperation } from "../internal/totals.js";

/** A bounded, one-pass grouping of rows already ordered into contiguous groups.
 * Every column needs a unique stable key; translated headings never select data or totals. */
export interface GroupedReportOptions<T = never> {
  readonly columns: readonly Column<T>[];
  readonly groupBy: string;
  /** Numeric metrics must declare type:"number" and resolve to finite numbers or null. Count ignores null cells. */
  readonly totals?: Readonly<Record<string, TotalOperation>>;
  /** Synchronous display text. The default is literal scalar text, an ISO UTC date, or empty text for null. */
  readonly groupLabel?: (value: CellValue) => string;
  /** Maximum distinct groups retained for recurrence detection. Defaults to 1,000. */
  readonly maxGroups?: number;
  /** Cumulative UTF-16 characters in retained group identifiers and captured labels. Defaults to 1,000,000. */
  readonly maxGroupKeyCharacters?: number;
  readonly signal?: AbortSignal;
  /** Source row, cell and text ceilings. Byte-output limits belong to the selected format writer. */
  readonly limits?: Pick<ExportLimits, "maxRows" | "maxCells" | "maxTextCharacters">;
}

export interface GroupedReportSummary {
  /** All accepted body rows, including rows with null or nonnumeric values. */
  readonly rows: number;
  readonly groups: number;
  /** Column-order totals. Columns without an operation have null totals. */
  readonly totals: readonly (number | null)[];
}

interface GroupIdentity {
  readonly groupIndex: number;
  readonly key: CellValue;
  readonly label: string;
}

/** Shared events consumed by grouped XLSX and PDF writers, without buffering a group's body. */
export type GroupedReportEvent =
  | ({ readonly kind: "groupStart" } & GroupIdentity)
  | { readonly kind: "row"; readonly values: readonly ExportValue[]; readonly rowIndex: number; readonly groupIndex: number }
  | ({ readonly kind: "groupEnd"; readonly rows: number; readonly totals: readonly (number | null)[] } & GroupIdentity)
  | ({ readonly kind: "complete" } & GroupedReportSummary);

export interface GroupedReportSource {
  /** Captured portable metadata in export order. Values have already been projected into positional rows. */
  readonly columns: readonly Column<readonly ExportValue[]>[];
  readonly groupBy: string;
  readonly totals: Readonly<Record<string, TotalOperation>>;
  /** Consume exactly once. Disposing or failing a writer returns the original row iterator. */
  readonly events: AsyncIterable<GroupedReportEvent>;
}

const sources = new WeakMap<object, { signal: AbortSignal | undefined; events(signal?: AbortSignal): AsyncIterable<GroupedReportEvent> }>();
/** @internal Writers accept the owned immutable event source rather than trusting caller-authored event protocols. */
export function assertGroupedReportSource(value: unknown): asserts value is GroupedReportSource {
  if (!value || typeof value !== "object" || !sources.has(value)) throw new TypeError("Create grouped report sources with createGroupedReport.");
}
/** @internal The consuming writer's signal reaches pending source reads as well as destination writes. */
export function groupedReportEvents(source: GroupedReportSource, signal?: AbortSignal): AsyncIterable<GroupedReportEvent> {
  assertGroupedReportSource(source);
  return { async *[Symbol.asyncIterator]() {
    const operation = groupedReportOperation(source, signal);
    try { yield* operation.events; } finally { operation.dispose(); }
  } };
}
/** @internal Share cancellation between original source reads and destination writes; callers dispose in finally. */
export function groupedReportOperation(source: GroupedReportSource, signal?: AbortSignal): {
  readonly signal: AbortSignal | undefined; readonly events: AsyncIterable<GroupedReportEvent>; dispose(): void
} {
  assertGroupedReportSource(source);
  const factory = sources.get(source)!, linked = linkedSignals(factory.signal, signal);
  return { signal: linked.signal, events: factory.events(linked.signal), dispose: linked.dispose };
}

function linkedSignals(source?: AbortSignal, consumer?: AbortSignal): { signal: AbortSignal | undefined; dispose(): void } {
  if (!source || !consumer || source === consumer) return { signal: source ?? consumer, dispose() {} };
  const controller = new AbortController(), abortSource = () => controller.abort(source.reason), abortConsumer = () => controller.abort(consumer.reason);
  source.addEventListener("abort", abortSource, { once: true }); consumer.addEventListener("abort", abortConsumer, { once: true });
  if (source.aborted) abortSource(); else if (consumer.aborted) abortConsumer();
  return { signal: controller.signal, dispose() { source.removeEventListener("abort", abortSource); consumer.removeEventListener("abort", abortConsumer); } };
}

function maximum(value: number | undefined, fallback: number, name: string): number {
  const result = value ?? fallback;
  if (!Number.isSafeInteger(result) || result < 0) throw new RangeError(name + " must be a nonnegative safe integer.");
  return result;
}

function copyValue(value: unknown): ExportValue {
  assertExportValue(value);
  if (isDate(value)) return new Date(Date.prototype.getTime.call(value));
  if (value instanceof ExportCell) {
    const scalar = value.value, text = value.text, presentation = value.presentation, link = value.link;
    return new ExportCell(isDate(scalar) ? new Date(Date.prototype.getTime.call(scalar)) : scalar,
      { ...(text === undefined ? {} : { text }), ...(presentation === undefined ? {} : { presentation }), ...(link === undefined ? {} : { link }) });
  }
  return value ?? null;
}

function scalar(value: ExportValue): CellValue { return value instanceof ExportCell ? value.value : value; }

function identity(value: CellValue): { key: CellValue; id: string; text: string } {
  if (value == null) return { key: null, id: "null", text: "" };
  if (isDate(value)) {
    const time = Date.prototype.getTime.call(value);
    if (!Number.isFinite(time)) throw new TypeError("A group key date must be valid.");
    return { key: new Date(time), id: "date:" + time, text: new Date(time).toISOString() };
  }
  if (typeof value === "number" && !Number.isFinite(value)) throw new TypeError("A numeric group key must be finite.");
  const text = String(value);
  return { key: value, id: typeof value + ":" + text, text };
}

function eventKey(value: CellValue): CellValue { return isDate(value) ? new Date(Date.prototype.getTime.call(value)) : value; }
function result(aggregates: readonly NumericAggregate[]): readonly (number | null)[] {
  return Object.freeze(aggregates.map(aggregate => aggregate.operation === undefined ? null : aggregate.value()));
}

/** Group a preordered row source once. Closed-group recurrence rejects rather than silently splitting a group.
 * Equality is locale independent: strings are exact, numbers and booleans retain their types, dates compare
 * by UTC instant, and null/undefined share one group. No sorting or complete group-body retention occurs. */
export function createGroupedReport<T>(rows: Iterable<T> | AsyncIterable<T>, options: GroupedReportOptions<T>): GroupedReportSource {
  if (!options || typeof options !== "object") throw new TypeError("Grouped report options are required.");
  // Capture configuration before validation so accessors cannot select different columns in later stages.
  const { columns: selected, groupBy, totals: selectedTotals, groupLabel, maxGroups: requestedGroups,
    maxGroupKeyCharacters: requestedCharacters, signal: sourceSignal, limits } = options;
  if (limits !== undefined && (!limits || typeof limits !== "object" || Array.isArray(limits)))
    throw new TypeError("Grouped source limits must be an object.");
  if (limits !== undefined && "maxOutputBytes" in limits)
    throw new TypeError("Set maxOutputBytes on the format writer, not the grouped source.");
  const capturedLimits = limits === undefined ? undefined : { ...limits };
  if (!Array.isArray(selected) || !selected.length) throw new TypeError("A grouped report needs declared columns.");
  const captured = selected.map(column => ({ ...column })) as ProjectionColumn[];
  const columns = copyColumns(captured);
  const keys = columns.map(column => column.key);
  if (keys.some(key => typeof key !== "string" || key.length === 0) || new Set(keys).size !== keys.length)
    throw new TypeError("Grouped report columns need unique, nonempty stable keys.");
  const groupColumn = keys.indexOf(groupBy);
  if (typeof groupBy !== "string" || groupColumn < 0) throw new TypeError("groupBy must name a declared stable column key.");
  if (groupLabel !== undefined && typeof groupLabel !== "function") throw new TypeError("groupLabel must be a synchronous function.");
  const totals = Object.freeze({ ...selectedTotals });
  const grand = createTotals(columns, totals);
  for (let column = 0; column < columns.length; column++)
    if (grand[column]!.operation !== undefined && columns[column]!.type !== "number")
      throw new TypeError("Grouped numeric totals require a column declared with type:number.");
  const maxGroups = maximum(requestedGroups, 1000, "maxGroups");
  const maxCharacters = maximum(requestedCharacters, 1000000, "maxGroupKeyCharacters");
  const budget = new ExportBudget(capturedLimits);
  const portable = Object.freeze(columns.map(column => {
    const { value: _getter, ...metadata } = column;
    return Object.freeze(metadata);
  })) as readonly Column<readonly ExportValue[]>[];
  for (const column of columns) {
    budget.cell(column.header);
    for (const group of column.groups ?? []) budget.reserve(0, group.length);
    Object.freeze(column);
  }
  let consumed = false;
  const eventsFor = (signal?: AbortSignal): AsyncIterable<GroupedReportEvent> => Object.freeze({ [Symbol.asyncIterator]() {
      if (consumed) throw new TypeError("A grouped report source can be consumed only once.");
      consumed = true;
      return iterate(signal);
    } });
  const events = eventsFor(sourceSignal);
  async function* iterate(signal?: AbortSignal): AsyncGenerator<GroupedReportEvent> {
    const project = createRowProjector(columns, undefined, signal);
    const seen = new Set<string>();
    let characters = 0, rowCount = 0, groups = 0, currentId: string | undefined;
    let key: CellValue = null, label = "", groupRows = 0, aggregates = createTotals(columns, totals);
    for await (const input of inputRows(rows, signal)) {
      // Copy positional input once before the shared projector validates and inspects its values.
      const snapshot = Array.isArray(input) ? Array.from(input) : input;
      const projected = project(snapshot, rowCount), values: ExportValue[] = [];
      for (let column = 0; column < columns.length; column++) values.push(copyValue(projected[column]));
      checkAbort(signal);
      const next = identity(scalar(values[groupColumn]));
      if (next.id !== currentId) {
        if (seen.has(next.id)) throw new TypeError("Grouped rows must be contiguous; a closed group key occurred again.");
        if (groups === maxGroups) throw new OfficeIMOError("RESOURCE_LIMIT", "maxGroups exceeded.");
        let nextLabel = next.text;
        if (groupLabel) {
          const formatted: unknown = groupLabel(eventKey(next.key));
          if (typeof (formatted as { then?: unknown } | null)?.then === "function") void Promise.resolve(formatted).catch(() => {});
          if (typeof formatted !== "string") throw new TypeError("groupLabel must return synchronous text.");
          nextLabel = formatted;
        }
        checkAbort(signal);
        const nextCharacters = characters + next.id.length + nextLabel.length;
        if (!Number.isSafeInteger(nextCharacters) || nextCharacters > maxCharacters)
          throw new OfficeIMOError("RESOURCE_LIMIT", "maxGroupKeyCharacters exceeded.");
        budget.reserve(0, nextLabel.length);
        if (currentId !== undefined) yield Object.freeze({ kind: "groupEnd", groupIndex: groups - 1,
          key: eventKey(key), label, rows: groupRows, totals: result(aggregates) });
        currentId = next.id; key = next.key; label = nextLabel; groupRows = 0;
        characters = nextCharacters; seen.add(next.id); groups++;
        aggregates = createTotals(columns, totals);
        yield Object.freeze({ kind: "groupStart", groupIndex: groups - 1, key: eventKey(key), label });
      }
      budget.row(rowCount + 1);
      for (let column = 0; column < values.length; column++) {
        const value = values[column], raw = scalar(value);
        if (grand[column]!.operation !== undefined && raw != null && (typeof raw !== "number" || !Number.isFinite(raw)))
          throw new TypeError("Grouped numeric metrics must resolve to finite numbers or null.");
        budget.cell(raw);
        if (value instanceof ExportCell && value.text !== undefined) budget.reserve(0, value.text.length);
        aggregates[column]!.accept(raw); grand[column]!.accept(raw);
      }
      checkAbort(signal);
      yield Object.freeze({ kind: "row", values: Object.freeze(values), rowIndex: rowCount, groupIndex: groups - 1 });
      rowCount++; groupRows++;
      if (taskYieldDue()) { await pause(); checkAbort(signal); }
    }
    checkAbort(signal);
    if (currentId !== undefined) yield Object.freeze({ kind: "groupEnd", groupIndex: groups - 1,
      key: eventKey(key), label, rows: groupRows, totals: result(aggregates) });
    checkAbort(signal);
    yield Object.freeze({ kind: "complete", rows: rowCount, groups, totals: result(grand) });
  }
  const source = Object.freeze({ columns: portable, groupBy, totals, events });
  sources.set(source, { signal: sourceSignal, events: eventsFor });
  return source;
}

import type { CellValue, Column, ExportLimits } from "./index.js";
import { ExportCell, assertScalar } from "./presentation.js";
import type { ExportValue } from "./presentation.js";
import { isDate } from "./dates.js";
import { checkAbort, inputRows, pause, taskYieldDue } from "./iteration.js";
import { ExportBudget } from "./limits.js";
import { OfficeIMOError } from "./errors.js";
import { copyColumns, createRowProjector } from "../internal/rows.js";
import type { ProjectionColumn } from "../internal/columns.js";
import { createTotals } from "../internal/totals.js";
import type { NumericAggregate, TotalOperation } from "../internal/totals.js";
import { copyExportValue } from "../internal/export-values.js";

export type { TotalOperation } from "../internal/totals.js";

/** Source-row position and stable grouping field, independent of visible export columns. */
export interface GroupedReportValueContext {
  readonly rowIndex: number;
  /** Zero-based grouping depth. */
  readonly level: number;
  readonly key: string;
}

/** One grouping level. Declared column keys reuse their projected value; hidden fields need a getter. */
export interface GroupedReportLevel<T = never> {
  readonly key: string;
  readonly value?: (row: T, context: GroupedReportValueContext) => CellValue;
  /** Synchronous display text; defaults to scalar text, an ISO UTC date, or empty text for null.
   * Identity always uses the typed value. */
  readonly label?: (value: CellValue) => string;
}

/** A bounded, one-pass hierarchy of rows already ordered into contiguous groups.
 * Every column needs a unique stable key; translated headings never select data or totals. */
export interface GroupedReportOptions<T = never> {
  readonly columns: readonly Column<T>[];
  /** One to four levels, outermost first. Hidden grouping fields never become output columns. */
  readonly groupBy: string | readonly (string | GroupedReportLevel<T>)[];
  /** Numeric metrics must declare type:"number" and resolve to finite numbers or null. Count ignores null cells. */
  readonly totals?: Readonly<Record<string, TotalOperation>>;
  /** Single-level shorthand. For a hierarchy, set each descriptor's label instead. */
  readonly groupLabel?: (value: CellValue) => string;
  /** Maximum opened groups across all depths, retained for recurrence detection. Defaults to 1,000. */
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
  /** All opened groups across every hierarchy level. */
  readonly groups: number;
  /** Column-order totals. Columns without an operation have null totals. */
  readonly totals: readonly (number | null)[];
}

/** A typed group within its parent. Labels are presentation and never select data or totals. */
export interface GroupedReportGroup {
  readonly groupIndex: number;
  readonly key: CellValue;
  readonly label: string;
  /** Zero-based grouping depth. */
  readonly level: number;
}

/** Shared events consumed by grouped XLSX and PDF writers, without buffering a group's body. */
export type GroupedReportEvent =
  | ({ readonly kind: "groupStart"; readonly path: readonly GroupedReportGroup[] } & GroupedReportGroup)
  | { readonly kind: "row"; readonly values: readonly ExportValue[]; readonly rowIndex: number;
      readonly groupIndex: number; readonly path: readonly GroupedReportGroup[] }
  | ({ readonly kind: "groupEnd"; readonly rows: number; readonly totals: readonly (number | null)[];
      readonly path: readonly GroupedReportGroup[] } & GroupedReportGroup)
  | ({ readonly kind: "complete" } & GroupedReportSummary);

declare const groupedReportSource: unique symbol;
export interface GroupedReportSource {
  /** Factory-owned source. Call createGroupedReport rather than implementing its event protocol. */
  readonly [groupedReportSource]: true;
  /** Captured portable metadata in export order. Values have already been projected into positional rows. */
  readonly columns: readonly Column<readonly ExportValue[]>[];
  /** Captured stable grouping keys, outermost first. */
  readonly groupBy: readonly string[];
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

interface CapturedLevel<T> {
  readonly key: string;
  readonly column: number;
  readonly value: GroupedReportLevel<T>["value"];
  readonly label: GroupedReportLevel<T>["label"];
}

function groupingLevels<T>(input: GroupedReportOptions<T>["groupBy"], keys: readonly (string | undefined)[],
  groupLabel: GroupedReportOptions<T>["groupLabel"]): readonly CapturedLevel<T>[] {
  const selected = typeof input === "string" ? [input] : Array.isArray(input) ? Array.from(input) : undefined;
  if (!selected || selected.length < 1 || selected.length > 4)
    throw new TypeError("groupBy must declare between one and four grouping levels.");
  if (groupLabel !== undefined && typeof groupLabel !== "function") throw new TypeError("groupLabel must be a synchronous function.");
  if (groupLabel && selected.length !== 1) throw new TypeError("groupLabel is only available for one grouping level; use descriptor labels for a hierarchy.");
  const seen = new Set<string>();
  return Object.freeze(selected.map(level => {
    if (typeof level !== "string" && (!level || typeof level !== "object" || Array.isArray(level)))
      throw new TypeError("A grouping level must name a stable column key or provide a descriptor.");
    // Supported structural fields, including prototype accessors, are captured once.
    const { key, value, label } = typeof level === "string" ? { key: level, value: undefined, label: undefined } : level;
    if (typeof key !== "string" || !key.length || seen.has(key)) throw new TypeError("Grouping levels need unique, nonempty stable keys.");
    seen.add(key);
    if (value !== undefined && typeof value !== "function") throw new TypeError("Grouping value getters must be synchronous functions.");
    if (label !== undefined && typeof label !== "function") throw new TypeError("Grouping labels must be synchronous functions.");
    if (label && groupLabel) throw new TypeError("Specify either groupLabel or the grouping descriptor label.");
    const column = keys.indexOf(key);
    if (column >= 0 && value !== undefined) throw new TypeError("A declared grouping column reuses its projected value; do not provide another value getter.");
    if (column < 0 && value === undefined) throw new TypeError("groupBy must name a declared stable column key; hidden grouping fields need an explicit value getter.");
    return Object.freeze({ key, column, value, label: label ?? groupLabel });
  }));
}

interface ActiveGroup {
  readonly id: string;
  readonly group: GroupedReportGroup;
  readonly aggregates: readonly NumericAggregate[];
  rows: number;
}

/** Copy date keys on every event so a consumer cannot mutate later hierarchy decisions. */
function groupPath(active: readonly ActiveGroup[]): readonly GroupedReportGroup[] {
  return Object.freeze(active.map(({ group }) => Object.freeze({ ...group, key: eventKey(group.key) })));
}

function labelText(format: GroupedReportLevel["label"], key: CellValue, fallback: string): string {
  if (!format) return fallback;
  const formatted: unknown = format(eventKey(key));
  if (typeof (formatted as { then?: unknown } | null)?.then === "function") void Promise.resolve(formatted).catch(() => {});
  if (typeof formatted !== "string") throw new TypeError("Grouping labels must return synchronous text.");
  return formatted;
}

/** Group a preordered row source once. Closed-group recurrence within the same parent rejects.
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
  if (!Array.isArray(selected) || !selected.length) throw new TypeError("A grouped report needs declared columns.");
  const columns = copyColumns(selected as readonly ProjectionColumn[]);
  const keys = columns.map(column => column.key);
  if (keys.some(key => typeof key !== "string" || key.length === 0) || new Set(keys).size !== keys.length)
    throw new TypeError("Grouped report columns need unique, nonempty stable keys.");
  const levels = groupingLevels(groupBy, keys, groupLabel);
  const totals = Object.freeze({ ...selectedTotals });
  const grand = createTotals(columns, totals);
  for (let column = 0; column < columns.length; column++)
    if (grand[column]!.operation !== undefined && columns[column]!.type !== "number")
      throw new TypeError("Grouped numeric totals require a column declared with type:number.");
  const maxGroups = maximum(requestedGroups, 1000, "maxGroups");
  const maxCharacters = maximum(requestedCharacters, 1000000, "maxGroupKeyCharacters");
  const budget = new ExportBudget(limits);
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
    const seen = new Set<string>(), active: ActiveGroup[] = [];
    let characters = 0, rowCount = 0, groups = 0;
    for await (const input of inputRows(rows, signal)) {
      // Copy positional input once before the shared projector validates and inspects its values.
      const snapshot = Array.isArray(input) ? Array.from(input) : input;
      const projected = project(snapshot, rowCount), values: ExportValue[] = [];
      for (let column = 0; column < columns.length; column++) values.push(copyExportValue(projected[column]) ?? null);
      checkAbort(signal);
      const next = levels.map((level, index) => {
        checkAbort(signal);
        const value = level.column >= 0 ? scalar(values[level.column]) : level.value!(snapshot as T,
          Object.freeze({ rowIndex: rowCount, level: index, key: level.key }));
        assertScalar(value); checkAbort(signal);
        return identity(value);
      });
      budget.row(rowCount + 1);
      for (let column = 0; column < values.length; column++) {
        const value = values[column], raw = scalar(value);
        if (grand[column]!.operation !== undefined && raw != null && (typeof raw !== "number" || !Number.isFinite(raw)))
          throw new TypeError("Grouped numeric metrics must resolve to finite numbers or null.");
        budget.cell(raw);
        if (value instanceof ExportCell && value.text !== undefined) budget.reserve(0, value.text.length);
      }
      let common = 0;
      while (common < active.length && active[common]!.id === next[common]!.id) common++;
      const pending: ActiveGroup[] = [];
      for (let level = common; level < levels.length; level++) {
        const value = next[level]!, parent = pending.at(-1) ?? active[level - 1];
        // Parent indexes are generated integers, so scalar strings containing delimiters cannot collide.
        const scopedId = (parent?.group.groupIndex ?? -1) + ":" + value.id;
        if (seen.has(scopedId)) throw new TypeError("Grouped rows must be contiguous; a closed group key occurred again within the same parent.");
        if (groups === maxGroups) throw new OfficeIMOError("RESOURCE_LIMIT", "maxGroups exceeded.");
        const label = labelText(levels[level]!.label, value.key, value.text);
        checkAbort(signal);
        const nextCharacters = characters + scopedId.length + label.length;
        if (!Number.isSafeInteger(nextCharacters) || nextCharacters > maxCharacters)
          throw new OfficeIMOError("RESOURCE_LIMIT", "maxGroupKeyCharacters exceeded.");
        budget.reserve(0, label.length);
        pending.push({ id: value.id, group: Object.freeze({ groupIndex: groups, key: value.key, label, level }),
          rows: 0, aggregates: createTotals(columns, totals) });
        characters = nextCharacters; seen.add(scopedId); groups++;
      }
      while (active.length > common) {
        checkAbort(signal);
        const path = groupPath(active), ending = active.pop()!;
        yield Object.freeze({ kind: "groupEnd", ...path.at(-1)!, path, rows: ending.rows, totals: result(ending.aggregates) });
      }
      for (const opening of pending) {
        checkAbort(signal); active.push(opening);
        const path = groupPath(active);
        yield Object.freeze({ kind: "groupStart", ...path.at(-1)!, path });
      }
      for (let column = 0; column < values.length; column++) {
        const raw = scalar(values[column]);
        for (const group of active) group.aggregates[column]!.accept(raw);
        grand[column]!.accept(raw);
      }
      for (const group of active) group.rows++;
      checkAbort(signal);
      yield Object.freeze({ kind: "row", values: Object.freeze(values), rowIndex: rowCount,
        groupIndex: active.at(-1)!.group.groupIndex, path: groupPath(active) });
      rowCount++;
      if (taskYieldDue()) { await pause(); checkAbort(signal); }
    }
    checkAbort(signal);
    while (active.length) {
      checkAbort(signal);
      const path = groupPath(active), ending = active.pop()!;
      yield Object.freeze({ kind: "groupEnd", ...path.at(-1)!, path, rows: ending.rows, totals: result(ending.aggregates) });
    }
    checkAbort(signal);
    yield Object.freeze({ kind: "complete", rows: rowCount, groups, totals: result(grand) });
  }
  const source = Object.freeze({ columns: portable, groupBy: Object.freeze(levels.map(level => level.key)), totals, events }) as GroupedReportSource;
  sources.set(source, { signal: sourceSignal, events: eventsFor });
  return source;
}

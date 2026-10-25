import { createGroupedReport } from "../../core/grouped-report.js";
import type { GroupedReportSource, GroupedReportLevel } from "../../core/grouped-report.js";
import { ExportCell } from "../../core/presentation.js";
import type { CellValue, ExportLimits, ExportValue } from "../../core/index.js";
import { checkAbort } from "../../core/iteration.js";
import { createTotals } from "../../internal/totals.js";
import type { TotalOperation } from "../../internal/totals.js";
import { captureFields } from "../../internal/capture.js";
import { datetime } from "./dates.js";
import type { CanopyCapture, CanopyDiagnostic, CanopyExport, CanopyExportGroup, CanopyFormat, CanopySourceOptions } from "./types.js";

interface Range {
  readonly offset: number;
  readonly end: number;
  readonly value: CellValue;
  readonly text: string;
  readonly expected: Readonly<Record<string, number | null>>;
}
const identity = (value: CellValue): string => value instanceof Date ? "date:" + value.getTime() : typeof value + ":" + String(value);
const scalar = (value: ExportValue): CellValue => value instanceof ExportCell ? value.value : value;

/** Validate captured ranges, then use the owned grouping factory and its existing writers.
 * Only ranges, labels and at most four active aggregate vectors are retained; no detail body is buffered. */
export function canopyGrouped(source: CanopyExport, request: CanopyCapture["request"], options: CanopySourceOptions,
  limits: ExportLimits | undefined, report: (diagnostic: CanopyDiagnostic) => void, format: CanopyFormat): GroupedReportSource | undefined {
  const input = request.grouping;
  if (input === undefined) return undefined; // query.groupBy never manufactures a group snapshot.
  if (request.values !== "raw") { report({ code: "GROUPING_REQUIRES_RAW_VALUES" }); return undefined; }
  if (source.rowCount === null) throw new TypeError("Captured grouping requires its exact scope recordCount.");
  const grouping = captureFields(input, ["levels", "aggregates", "totals", "groups"]);
  if (!Array.isArray(grouping.levels) || grouping.levels.length < 1 || grouping.levels.length > 4 || !Array.isArray(grouping.aggregates) || !Array.isArray(grouping.groups))
    throw new TypeError("Canopy grouping needs one through four levels, aggregates and ordered groups.");
  const levels = grouping.levels.map(level => Object.freeze(captureFields(level, ["column", "direction", "title"])));
  if (levels.some(level => typeof level.column !== "string" || !level.column || level.direction !== "asc" && level.direction !== "desc") ||
      new Set(levels.map(level => level.column)).size !== levels.length) throw new TypeError("Canopy grouping levels need unique stable columns and directions.");
  const declarations = grouping.aggregates.map(aggregate => Object.freeze(captureFields(aggregate, ["id", "column", "operation", "label"])));
  const ids = new Set<string>(), totals: Record<string, TotalOperation> = Object.create(null);
  const operations: Readonly<Record<string, TotalOperation>> = { sum: "sum", count: "count", average: "average", minimum: "min", maximum: "max" };
  for (const aggregate of declarations) {
    if (typeof aggregate.id !== "string" || !aggregate.id || ids.has(aggregate.id)) throw new TypeError("Canopy aggregate IDs must be unique nonempty strings.");
    ids.add(aggregate.id);
    const column = source.columns.find(column => column.key === aggregate.column);
    if (!column || column.type !== "number") throw new TypeError("Canopy aggregate declarations require an exported raw numeric column.");
    if (Object.hasOwn(totals, aggregate.column)) throw new TypeError("The grouped report supports one aggregate operation per column.");
    const operation = operations[aggregate.operation];
    if (operation === undefined) { report({ code: "UNSUPPORTED_GROUP_AGGREGATE:" + aggregate.operation, columnId: aggregate.column }); return undefined; }
    totals[aggregate.column] = operation;
  }
  // Reuse the canonical operation validator rather than implementing aggregate semantics here.
  createTotals(source.columns, totals);
  const expected = (input: Readonly<Record<string, number | null>>): Readonly<Record<string, number | null>> => {
    if (!input || typeof input !== "object" || Array.isArray(input)) throw new TypeError("Canopy aggregate values must be an ID-keyed object.");
    const snapshot = { ...input };
    if (Object.keys(snapshot).some(key => !ids.has(key)) || declarations.some(aggregate => !Object.hasOwn(snapshot, aggregate.id)))
      throw new TypeError("Canopy aggregate values must reference exactly the captured declarations.");
    for (const value of Object.values(snapshot)) if (value !== null && (typeof value !== "number" || !Number.isFinite(value)))
      throw new TypeError("Canopy aggregate values must be finite numbers or null.");
    return Object.freeze(snapshot);
  };
  const grandExpected = expected(grouping.totals), byDepth: Range[][] = levels.map(() => []), labels = levels.map(() => new Map<string, string>());
  let groups = 0, characters = 0;
  const ranges = (inputs: readonly CanopyExportGroup[], offset: number, end: number, depth: number): void => {
    if (!Array.isArray(inputs)) throw new TypeError("Canopy groups must be ordered arrays.");
    let cursor = offset;
    const keys = new Set<string>(), values = new Set<string>();
    for (const input of inputs) {
      const group = captureFields(input, ["key", "value", "text", "offset", "count", "aggregates", "groups"]);
      if (!group || typeof group.key !== "string" || keys.has(group.key) || !Number.isSafeInteger(group.offset) || !Number.isSafeInteger(group.count) ||
          group.count <= 0 || group.offset !== cursor || group.count > end - group.offset)
        throw new TypeError("Canopy group ranges must be nonempty contiguous half-open partitions within their parent.");
      if (group.value !== null && !["string", "number", "boolean"].includes(typeof group.value) ||
          typeof group.value === "number" && !Number.isFinite(group.value))
        throw new TypeError("Canopy group values must be portable finite scalars.");
      const dateColumn = request.columns.find(column => column.id === levels[depth]!.column)?.kind === "datetime";
      const value = format === "xlsx" && dateColumn && typeof group.value === "string"
        ? datetime(group.value, options.datetime ?? "preserve", options.xlsx?.dateMode ?? request.timeZone) : group.value;
      const key = identity(value), text = group.text ?? String(group.value ?? "");
      if (typeof text !== "string" || values.has(key)) throw new TypeError("A Canopy group key cannot recur within the same parent.");
      const previousLabel = labels[depth]!.get(key);
      if (previousLabel !== undefined && previousLabel !== text) throw new TypeError("The same grouping value needs consistent captured text at one level.");
      labels[depth]!.set(key, text); keys.add(group.key); values.add(key);
      if (++groups > 1000 || (characters += group.key.length + key.length + text.length) > 1000000)
        throw new RangeError("Canopy group metadata exceeds the existing bounded report profile.");
      const next = group.offset + group.count;
      byDepth[depth]!.push(Object.freeze({ offset: group.offset, end: next, value, text, expected: expected(group.aggregates) }));
      if (depth + 1 < levels.length) ranges(group.groups ?? [], group.offset, next, depth + 1);
      else if (group.groups !== undefined && group.groups.length) throw new TypeError("Canopy nested groups exceed their declared levels.");
      cursor = next;
    }
    if (cursor !== end) throw new TypeError("Canopy grouping does not cover its complete exported detail range.");
  };
  ranges(grouping.groups, 0, source.rowCount, 0);
  // Validate the producer's wire arithmetic separately from the grouped writer's native formula caches.
  // Canopy sum/average use compensated accumulation and return null without numeric inputs.
  const wireTotals = () => {
    const ordinary = createTotals(source.columns, totals), sums = source.columns.map(() => 0), corrections = source.columns.map(() => 0), counts = source.columns.map(() => 0);
    return {
      accept(values: readonly ExportValue[]): void {
        for (let index = 0; index < values.length; index++) {
          const value = scalar(values[index]), operation = totals[source.columns[index]!.key];
          if (operation !== "sum" && operation !== "average") { ordinary[index]!.accept(value); continue; }
          if (typeof value !== "number" || !Number.isFinite(value)) continue;
          counts[index] = counts[index]! + 1;
          const next = value - corrections[index]!, total = sums[index]! + next;
          corrections[index] = total - sums[index]! - next; sums[index] = total;
        }
      },
      value(index: number): number | null {
        const operation = totals[source.columns[index]!.key];
        return operation === "sum" || operation === "average" ? !counts[index] ? null
          : operation === "sum" ? sums[index]! : sums[index]! / counts[index]! : ordinary[index]!.value();
      }
    };
  };
  const indexes = levels.map(() => 0), active: ({ range: Range; aggregates: ReturnType<typeof wireTotals> } | undefined)[] = levels.map(() => undefined);
  const at = (depth: number, row: number): Range => {
    const list = byDepth[depth]!;
    while (list[indexes[depth]!] && list[indexes[depth]!]!.end <= row) indexes[depth] = indexes[depth]! + 1;
    const range = list[indexes[depth]!]!;
    if (!range || row < range.offset || row >= range.end) throw new TypeError("A detail row is outside its captured Canopy grouping.");
    return range;
  };
  const verify = (actual: ReturnType<typeof wireTotals>, expected: Readonly<Record<string, number | null>>): void => {
    for (const declaration of declarations) {
      const position = source.columns.findIndex(column => column.key === declaration.column), observed = actual.value(position), declared = expected[declaration.id];
      // Exact producer-wire comparison retains wrong-scope rejection without rounding native caches or applying an epsilon.
      if (observed !== declared) throw new TypeError("Canopy captured aggregate differs from its consumed rows: " + declaration.id + ".");
    }
  };
  const same = (actual: CellValue, expected: CellValue): boolean => actual instanceof Date && expected instanceof Date
    ? actual.getTime() === expected.getTime() : actual === expected;
  async function* rows(): AsyncGenerator<readonly ExportValue[]> {
    let index = 0;
    const grand = wireTotals();
    for await (const values of source.rows()) {
      checkAbort(options.signal);
      for (let depth = 0; depth < levels.length; depth++) {
        const range = at(depth, index), column = source.columns.findIndex(column => column.key === levels[depth]!.column);
        if (column >= 0 && !same(scalar(values[column]), range.value)) throw new TypeError("Canopy detail value disagrees with its captured group range.");
        if (range.offset === index) active[depth] = { range, aggregates: wireTotals() };
      }
      grand.accept(values);
      for (const group of active) group!.aggregates.accept(values);
      yield values; index++;
      for (let depth = levels.length - 1; depth >= 0; depth--) {
        const group = active[depth]!;
        if (group.range.end === index) { verify(group.aggregates, group.range.expected); active[depth] = undefined; }
      }
    }
    verify(grand, grandExpected);
  }
  const groupBy: GroupedReportLevel<readonly ExportValue[]>[] = levels.map((level, depth) => {
    const visible = source.columns.some(column => column.key === level.column);
    return { key: level.column,
      ...(visible ? {} : { value: (_row: readonly ExportValue[], context: { rowIndex: number }) => at(depth, context.rowIndex).value }),
      label: (value: CellValue) => labels[depth]!.get(identity(value)) ?? String(value ?? "") };
  });
  const sourceLimits = limits === undefined ? undefined : {
    ...(limits.maxRows === undefined ? {} : { maxRows: limits.maxRows }),
    ...(limits.maxCells === undefined ? {} : { maxCells: limits.maxCells }),
    ...(limits.maxTextCharacters === undefined ? {} : { maxTextCharacters: limits.maxTextCharacters })
  };
  return createGroupedReport(rows(), { columns: source.columns, groupBy, totals,
    ...(options.signal === undefined ? {} : { signal: options.signal }), ...(sourceLimits === undefined ? {} : { limits: sourceLimits }) });
}

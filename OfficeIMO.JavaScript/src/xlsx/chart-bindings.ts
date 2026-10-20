import { checkAbort } from "../core/iteration.js";
import { NotSupportedError, OfficeIMOError } from "../core/errors.js";
import { cleanXml } from "../xml/index.js";
import type { InvalidCharacterPolicy } from "../xml/index.js";
import type { ChartColumn, ChartData, ChartRangeData, ChartRangeSeries, ChartType } from "./chart-types.js";

/** @internal Workbook-owned identity, column resolution and finite A1 references. */
export interface ResolvedChartBinding {
  readonly source: object;
  readonly sheetName: string;
  readonly headerRows: number;
  readonly categoryColumn: number;
  readonly seriesColumns: readonly number[];
  readonly seriesNames: readonly string[];
  readonly categoryFormula: string;
  readonly valueFormulas: readonly string[];
  readonly nameFormulas: readonly (string | undefined)[];
}

/** @internal Dynamic category text shares the workbook's cumulative chart budget. */
export interface ChartBindingBudgets {
  readonly availablePoints: number;
  readonly availableText: number;
  readonly retainText: (characters: number) => void;
  readonly policy: InvalidCharacterPolicy;
}

/** @internal References and immutable caches are serialized by the ordinary chart writer. */
export interface PreparedChartBinding {
  readonly source: object;
  readonly points: number;
  readonly textCharacters: number;
  readonly data: ChartData;
  readonly references: {
    readonly categories: string;
    readonly values: readonly string[];
    readonly names: readonly (string | undefined)[];
  };
  observe(dataRowIndex: number, column: number, value: unknown): void;
  finish(): void;
}

function record(value: unknown, allowed: readonly string[], context: string): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError(context + " must be an object.");
  for (const key of Object.keys(value)) if (!allowed.includes(key)) throw new NotSupportedError(context + "." + key);
  return { ...value };
}

function column(value: unknown): asserts value is ChartColumn {
  if (typeof value === "string" ? !value.length : typeof value !== "number" || !Number.isInteger(value) || value < 1 || value > 16384)
    throw new TypeError("Chart columns require a declared key or a one-based column number.");
}

/** @internal Capture each caller property once before resolving workbook identities. */
export function captureChartRange(input: unknown): ChartRangeData {
  const value = record(input, ["kind", "sheet", "firstRow", "rowCount", "categoryColumn", "series"], "chart.data");
  if (value.kind !== "range") throw new TypeError("A chart range must declare kind: range.");
  if (!value.sheet || typeof value.sheet !== "object" || Array.isArray(value.sheet)) throw new TypeError("A chart range needs its owning worksheet.");
  const firstRow = value.firstRow === undefined ? 0 : value.firstRow, rowCount = value.rowCount;
  if (!Number.isSafeInteger(firstRow) || (firstRow as number) < 0 || !Number.isSafeInteger(rowCount) || (rowCount as number) < 1 ||
    (firstRow as number) + (rowCount as number) > 1048576) throw new RangeError("Chart data rows must form a nonempty finite Excel range.");
  column(value.categoryColumn);
  if (!Array.isArray(value.series) || value.series.length < 1 || value.series.length > 255) throw new RangeError("A chart needs from 1 through 255 series.");
  const series = value.series.map(item => {
    const copy = record(item, ["column", "name", "color"], "chart.data.series"); column(copy.column);
    if (copy.name !== undefined && typeof copy.name !== "string") throw new TypeError("Chart series name must be a string.");
    if (copy.color !== undefined && (typeof copy.color !== "string" || !/^#?[0-9a-fA-F]{6}$/.test(copy.color))) throw new TypeError("Chart colors must be opaque six-digit RGB hex values.");
    return Object.freeze(copy) as unknown as ChartRangeSeries;
  });
  return Object.freeze({ kind: "range", sheet: value.sheet, firstRow, rowCount, categoryColumn: value.categoryColumn, series: Object.freeze(series) }) as ChartRangeData;
}

/**
 * @internal Retain only the declared finite cache while the owning sheet writes once.
 * Initial series-name text is returned for registration; category text is charged
 * before copying through retainText. The workbook resolves source identity and A1.
 */
export function prepareChartBinding(input: ChartRangeData, descriptor: ResolvedChartBinding, budgets: ChartBindingBudgets,
  signal?: AbortSignal, type: ChartType = "columnClustered"): PreparedChartBinding {
  checkAbort(signal);
  const count = input.rowCount, first = input.firstRow ?? 0, points = count * input.series.length, categoryColumn = descriptor.categoryColumn;
  if (points > budgets.availablePoints) throw new OfficeIMOError("RESOURCE_LIMIT", "maxChartPoints exceeded.");
  if (descriptor.source !== input.sheet || !Number.isInteger(descriptor.headerRows) || descriptor.headerRows < 0 || descriptor.headerRows + first + count > 1048576 ||
    descriptor.seriesColumns.length !== input.series.length || descriptor.seriesNames.length !== input.series.length ||
    descriptor.valueFormulas.length !== input.series.length || descriptor.nameFormulas.length !== input.series.length)
    throw new TypeError("Resolved chart range does not match its declared columns and data rows.");
  for (const index of [descriptor.categoryColumn, ...descriptor.seriesColumns])
    if (!Number.isInteger(index) || index < 1 || index > 16384) throw new RangeError("Resolved chart columns must be one-based Excel columns.");
  if (descriptor.seriesColumns.includes(categoryColumn)) throw new TypeError("Chart category and numeric series must use different columns.");
  let textCharacters = 0;
  const names = input.series.map((series, i) => {
    const name = series.name ?? descriptor.seriesNames[i]!;
    if (typeof name !== "string" || name.length > 32767) throw new TypeError("Chart series names must fit Excel text cells.");
    textCharacters += name.length;
    if (textCharacters > budgets.availableText) throw new OfficeIMOError("RESOURCE_LIMIT", "maxChartTextCharacters exceeded.");
    return cleanXml(name, budgets.policy);
  });
  const colors = input.series.map(series => series.color?.replace(/^#/, "").toUpperCase());
  const categories = new Array<string>(count), values = input.series.map(() => new Array<number>(count));
  const captured = input.series.map(() => 0), positive = input.series.map(() => false); let categoryCount = 0, completed: ChartData | undefined;
  const finish = (): void => {
    checkAbort(signal);
    if (completed) return;
    if (categoryCount !== count || captured.some(value => value !== count)) throw new TypeError("Chart range did not receive every declared data row and numeric value.");
    if (type === "pie" && !positive[0]) throw new RangeError("A pie chart needs at least one positive value.");
    completed = Object.freeze({ categories: Object.freeze(categories), series: Object.freeze(values.map((points, i) => Object.freeze({
      name: names[i]!, values: Object.freeze(points), ...(colors[i] === undefined ? {} : { color: colors[i]! })
    }))) });
  };
  const data: ChartData = Object.freeze({ get categories() { finish(); return completed!.categories; }, get series() { finish(); return completed!.series; } });
  const seriesByColumn = new Map<number, number[]>();
  descriptor.seriesColumns.forEach((index, i) => { const indexes = seriesByColumn.get(index) ?? []; indexes.push(i); seriesByColumn.set(index, indexes); });
  return Object.freeze({ source: descriptor.source, points, textCharacters, data,
    references: Object.freeze({ categories: descriptor.categoryFormula, values: Object.freeze([...descriptor.valueFormulas]),
      names: Object.freeze(descriptor.nameFormulas.map((formula, i) => input.series[i]!.name === undefined ? formula : undefined)) }),
    observe(dataRowIndex: number, columnIndex: number, value: unknown): void {
      checkAbort(signal);
      const position = dataRowIndex - first;
      if (position < 0 || position >= count) return;
      if (completed) throw new OfficeIMOError("INVALID_STATE", "Chart cache is already finalized.");
      if (columnIndex === categoryColumn) {
        if (typeof value !== "string" || value.length > 32767) throw new TypeError("Chart category cells must contain strings of at most 32,767 UTF-16 units.");
        if (categories[position] !== undefined) throw new OfficeIMOError("INVALID_STATE", "A chart category cell was observed more than once.");
        budgets.retainText(value.length); categories[position] = cleanXml(value, budgets.policy); categoryCount++;
      }
      for (const i of seriesByColumn.get(columnIndex) ?? []) {
        if (typeof value !== "number" || !Number.isFinite(value)) throw new TypeError("Chart series cells require finite numbers or numeric formula caches.");
        if (type === "pie" && value < 0) throw new RangeError("Pie charts require nonnegative values.");
        if (values[i]![position] !== undefined) throw new OfficeIMOError("INVALID_STATE", "A chart value cell was observed more than once.");
        values[i]![position] = value; captured[i] = captured[i]! + 1; if (value > 0) positive[i] = true;
      }
    }, finish
  });
}

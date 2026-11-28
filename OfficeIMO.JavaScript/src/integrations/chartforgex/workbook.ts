import { ExportCell } from "../../core/presentation.js";
import { ExportBudget } from "../../core/limits.js";
import type { ExportLimits } from "../../core/limits.js";
import type { ColumnSettings, ExportValue } from "../../core/index.js";
import type { PreparedChart, PreparedPoint } from "./capture.js";

export type ChartColumn = ColumnSettings & { readonly key: string };
export interface ChartTable {
  readonly columns: readonly ChartColumn[];
  readonly rows: Iterable<readonly ExportValue[]>;
}

const display = (header: string): string => header + " (display)";
function dateFormat(chart: PreparedChart): string {
  let time = false, seconds = false;
  for (const series of chart.series) for (const point of series.points) {
    if (!(point.x instanceof Date)) continue;
    const within = point.x.getTime() % 86400000;
    if (within) time = true;
    if (within % 60000) seconds = true;
  }
  return seconds ? "yyyy-mm-dd hh:mm:ss" : time ? "yyyy-mm-dd hh:mm" : "yyyy-mm-dd";
}
function xColumn(chart: PreparedChart): ChartColumn {
  const header = chart.xTitle;
  if (chart.kind === "date") return { key: "x", header, type: "date", format: chart.xFormat ?? dateFormat(chart) };
  if (chart.kind === "number") return { key: "x", header, type: "number", ...(chart.xFormat === undefined ? {} : { format: chart.xFormat }) };
  return { key: "x", header, type: "string" };
}
const withUnit = (header: string, unit: string | undefined): string => unit ? header + " (" + unit + ")" : header;

/** Whether every series can share one row per x value. */
function canWiden(chart: PreparedChart): boolean {
  const captions = new Map<PreparedPoint["key"], string>();
  for (const series of chart.series) {
    const seen = new Set<PreparedPoint["key"]>();
    for (const point of series.points) {
      if (point.key === null && chart.kind !== "category") return false;
      if (seen.has(point.key)) return false;
      seen.add(point.key);
      if (point.xCaption !== undefined) {
        const previous = captions.get(point.key);
        if (previous !== undefined && previous !== point.xCaption) return false;
        captions.set(point.key, point.xCaption);
      }
    }
  }
  return true;
}

/** Map the chart's series to typed columns in a stable order: x, its caption, then each series' value and caption. */
export function chartTable(chart: PreparedChart, layout: "auto" | "wide" | "long" = "auto", limits?: ExportLimits): ChartTable {
  if (layout !== "auto" && layout !== "wide" && layout !== "long") throw new TypeError("Chart layout must be auto, wide or long.");
  const wide = layout !== "long" && canWiden(chart);
  if (layout === "wide" && !wide) throw new TypeError("A wide chart sheet needs unique x values in every series and consistent x captions.");
  const budget = new ExportBudget(limits);
  const headingRows = chart.title ? 2 : 1;
  function checkRows(count: number): void {
    budget.row(count);
    if (count + headingRows > 1048576) throw new RangeError("Excel supports at most 1,048,576 rows including headings.");
  }
  function checkCells(rows: number, columns: number): void {
    budget.check("maxCells", (rows + headingRows) * columns);
  }
  const x = xColumn(chart), hasXCaption = chart.series.some(series => series.points.some(point => point.xCaption !== undefined));
  const value = (point: PreparedPoint | undefined, format: string | undefined, column: string | undefined): ExportValue =>
    point?.y === null || point === undefined ? null : format !== undefined && format !== column ? new ExportCell(point.y, { presentation: { numberFormat: format } }) : point.y;
  if (!wide) {
    const yFormat = chart.yFormat, hasCaption = chart.series.some(series => series.points.some(point => point.caption !== undefined));
    const units = new Set(chart.series.filter(series => series.points.length > 0).map(series => (series.unit ?? chart.yUnit) || undefined));
    const mixedUnits = units.size > 1;
    const yHeader = withUnit(chart.yTitle, units.size === 0 ? chart.yUnit : mixedUnits ? undefined : units.values().next().value);
    const columns: ChartColumn[] = [{ key: "series", header: "Series", type: "string" }, { key: "seriesId", header: "Series ID", type: "string" }, x,
      ...(hasXCaption ? [{ key: "xCaption", header: display(x.header), type: "string" }] : []),
      { key: "y", header: yHeader, type: "number", ...(yFormat === undefined ? {} : { format: yFormat }) },
      ...(hasCaption ? [{ key: "caption", header: display(yHeader), type: "string" }] : []),
      ...(mixedUnits ? [{ key: "unit", header: "Unit", type: "string" }] : [])];
    const rowCount = chart.series.reduce((count, series) => count + series.points.length, 0);
    checkRows(rowCount);
    checkCells(rowCount, columns.length);
    function* rows(): Iterable<readonly ExportValue[]> {
      for (const series of chart.series) for (const point of series.points)
        yield [series.name, series.id, point.x, ...(hasXCaption ? [point.xCaption ?? null] : []), value(point, series.format ?? yFormat, yFormat),
          ...(hasCaption ? [point.caption ?? null] : []), ...(mixedUnits ? [series.unit ?? chart.yUnit ?? null] : [])];
    }
    return { columns, rows: rows() };
  }
  const captionColumns = chart.series.map(series => series.points.some(point => point.caption !== undefined));
  const columnCount = 1 + Number(hasXCaption) + chart.series.length + captionColumns.filter(Boolean).length;
  if (columnCount > 16384) throw new RangeError("Excel supports at most 16,384 columns.");
  const keys: (number | string | null)[] = [], captions = new Map<number | string | null, string>(), values = new Map<number | string | null, Date | number | string | null>();
  const known = new Set<number | string | null>();
  for (const series of chart.series) for (const point of series.points) {
    if (!known.has(point.key)) { checkRows(keys.length + 1); known.add(point.key); keys.push(point.key); values.set(point.key, point.x); }
    if (point.xCaption !== undefined && !captions.has(point.key)) captions.set(point.key, point.xCaption);
  }
  checkCells(keys.length, columnCount);
  if (chart.kind !== "category") keys.sort((a, b) => (a as number) - (b as number));
  const names = new Map<string, number>();
  for (const series of chart.series) names.set(series.name, (names.get(series.name) ?? 0) + 1);
  const columns: ChartColumn[] = [x, ...(hasXCaption ? [{ key: "xCaption", header: display(x.header), type: "string" }] : [])];
  const lookups = chart.series.map((series, index) => {
    const format = series.format ?? chart.yFormat, header = withUnit(series.name + (names.get(series.name)! > 1 ? " [" + series.id + "]" : ""), series.unit ?? chart.yUnit);
    const caption = captionColumns[index];
    columns.push({ key: "s" + index, header, type: "number", ...(format === undefined ? {} : { format }) });
    if (caption) columns.push({ key: "s" + index + "Caption", header: display(header), type: "string" });
    return { caption, format, points: new Map(series.points.map(point => [point.key, point])) };
  });
  function* rows(): Iterable<readonly ExportValue[]> {
    for (const key of keys) {
      const row: ExportValue[] = [values.get(key) ?? null, ...(hasXCaption ? [captions.get(key) ?? null] : [])];
      for (const lookup of lookups) {
        const point = lookup.points.get(key);
        row.push(point?.y ?? null);
        if (lookup.caption) row.push(point?.caption ?? null);
      }
      yield row;
    }
  }
  return { columns, rows: rows() };
}

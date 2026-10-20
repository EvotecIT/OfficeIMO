import { checkAbort } from "../core/iteration.js";
import { NotSupportedError, OfficeIMOError } from "../core/errors.js";
import { cleanXml } from "../xml/index.js";
import type { InvalidCharacterPolicy } from "../xml/index.js";
import { copyDrawingAnchor } from "./attachments.js";
import { captureFields } from "../internal/capture.js";
import { captureChartRange, prepareChartBinding } from "./chart-bindings.js";
import type { PreparedChartBinding, ResolvedChartBinding } from "./chart-bindings.js";
import type { WorksheetChart, ChartData, ChartRangeData, ChartDataLabels, ChartType } from "./chart-types.js";

/** @internal Captured chart data never retains the caller's arrays or options. */
export interface PreparedChart extends Omit<WorksheetChart, "data"> {
  readonly data: ChartData;
  readonly binding?: PreparedChartBinding;
  readonly points: number;
  readonly textCharacters: number;
}
/** @internal Workbook-unique identities bind the drawing, chart and source worksheet. */
export interface RegisteredChart extends PreparedChart { readonly id: number; readonly dataSheetName: string; }
/** @internal Workbook checks identity/lifecycle, resolves its range and owns the cumulative text reservation. */
export type ResolveChartBinding = (input: ChartRangeData) => {
  readonly descriptor: ResolvedChartBinding;
  readonly retainText: (characters: number) => void;
};

function fields(value: unknown, allowed: readonly string[], context: string): asserts value is Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError(context + " must be an object.");
  for (const key of Object.keys(value)) if (!allowed.includes(key)) throw new NotSupportedError(context + "." + key);
}

function snapshot<T extends object>(value: T, allowed: readonly string[], context: string): T {
  fields(value, allowed, context);
  return captureFields(value, allowed) as T;
}

/** @internal Preflight bounds before copying any chart data. Failed registration reserves nothing. */
export function prepareChart(input: WorksheetChart, policy: InvalidCharacterPolicy, availablePoints: number, availableText: number,
  signal?: AbortSignal, resolveBinding?: ResolveChartBinding): PreparedChart {
  checkAbort(signal);
  input = snapshot(input, ["type", "data", "row", "column", "width", "height", "description", "title", "language", "legend", "dataLabels", "categoryAxis", "valueAxis"], "chart");
  if (!["columnClustered", "barClustered", "line", "pie"].includes(input.type)) throw new NotSupportedError("chart type " + input.type);
  const type: ChartType = input.type;
  const requested = snapshot(input.data as unknown as Record<string, unknown>, ["kind", "sheet", "firstRow", "rowCount", "categoryColumn", "categories", "series"], "chart.data");
  const bound = requested.kind === "range" ? captureChartRange(requested) : undefined;
  if (!bound) fields(requested, ["categories", "series"], "chart.data");
  const categories = bound ? undefined : requested.categories, series = bound?.series ?? requested.series;
  if (!bound && !Array.isArray(categories)) throw new TypeError("Chart categories must be a nonempty array of strings.");
  if (!Array.isArray(series)) throw new RangeError("A chart needs from 1 through 255 series.");
  const categoryCount = bound?.rowCount ?? (categories as readonly unknown[]).length, seriesCount = series.length;
  if (!categoryCount) throw new TypeError("Chart categories must be a nonempty array of strings.");
  if (!seriesCount || seriesCount > 255) throw new RangeError("A chart needs from 1 through 255 series.");
  if (!bound && categoryCount > 1048575) throw new RangeError("Chart data must fit Excel's rows including its header.");
  if (type === "pie" && seriesCount !== 1) throw new TypeError("A pie chart supports exactly one series.");
  const points = categoryCount * seriesCount;
  if (points > availablePoints) throw new OfficeIMOError("RESOURCE_LIMIT", "maxChartPoints exceeded.");
  let textCharacters = 0;
  const text = (value: unknown, context: string): string => {
    if (typeof value !== "string") throw new TypeError(context + " must be a string.");
    if (value.length > 32767) throw new RangeError(context + " exceeds 32,767 UTF-16 units.");
    textCharacters += value.length;
    if (textCharacters > availableText) throw new OfficeIMOError("RESOURCE_LIMIT", "maxChartTextCharacters exceeded.");
    return cleanXml(value, policy);
  };
  const optionalText = (value: unknown, context: string) => value === undefined ? undefined : text(value, context);
  const title = optionalText(input.title, "Chart title"), description = optionalText(input.description, "Chart description");
  const requestedLanguage = optionalText(input.language, "Chart language");
  let language: string | undefined;
  if (requestedLanguage !== undefined) {
    try { language = Intl.getCanonicalLocales(requestedLanguage)[0]; if (!language) throw new RangeError(); }
    catch { throw new TypeError("Chart language must be a language tag such as pl or pl-PL."); }
  }
  const anchor = copyDrawingAnchor({ row: input.row, column: input.column, width: input.width, height: input.height,
    ...(description === undefined ? {} : { description }) }, policy);
  if (input.legend !== undefined && input.legend !== false && !["right", "bottom", "left", "top"].includes(input.legend)) throw new TypeError("Unsupported chart legend position.");
  let dataLabels: ChartDataLabels | undefined;
  if (input.dataLabels !== undefined) {
    const labels = snapshot(input.dataLabels, ["showValue", "showCategoryName", "showSeriesName", "showPercent", "position", "numberFormat"], "chart.dataLabels");
    for (const key of ["showValue", "showCategoryName", "showSeriesName", "showPercent"] as const)
      if (labels[key] !== undefined && typeof labels[key] !== "boolean") throw new TypeError("Chart label flags must be booleans.");
    if (labels.showPercent && type !== "pie") throw new TypeError("Percent data labels require a pie chart.");
    const positions = type === "pie" ? ["center", "insideEnd", "outsideEnd", "bestFit"] : type === "line" ? ["center", "above", "below", "left", "right"] : ["center", "insideEnd", "outsideEnd"];
    if (labels.position !== undefined && !positions.includes(labels.position)) throw new TypeError("Data label position is unsupported for this chart type.");
    const numberFormat = optionalText(labels.numberFormat, "Data label number format");
    if (numberFormat !== undefined && !numberFormat.trim()) throw new TypeError("Data label number format cannot be blank.");
    dataLabels = Object.freeze({ ...labels, ...(numberFormat === undefined ? {} : { numberFormat }) });
  }
  if (type === "pie" && (input.categoryAxis !== undefined || input.valueAxis !== undefined)) throw new TypeError("Pie charts have no axes.");
  let categoryAxis: WorksheetChart["categoryAxis"], valueAxis: WorksheetChart["valueAxis"];
  if (input.categoryAxis !== undefined) {
    const axisTitle = optionalText(snapshot(input.categoryAxis, ["title"], "chart.categoryAxis").title, "Category axis title");
    categoryAxis = Object.freeze(axisTitle === undefined ? {} : { title: axisTitle });
  }
  if (input.valueAxis !== undefined) {
    const axis = snapshot(input.valueAxis, ["title", "minimum", "maximum", "numberFormat"], "chart.valueAxis");
    for (const value of [axis.minimum, axis.maximum]) if (value !== undefined && !Number.isFinite(value)) throw new TypeError("Chart axis bounds must be finite numbers.");
    if (axis.minimum !== undefined && axis.maximum !== undefined && axis.minimum >= axis.maximum) throw new RangeError("Chart axis minimum must be below its maximum.");
    const axisTitle = optionalText(axis.title, "Value axis title"), numberFormat = optionalText(axis.numberFormat, "Value axis number format");
    if (numberFormat !== undefined && !numberFormat.trim()) throw new TypeError("Value axis number format cannot be blank.");
    valueAxis = Object.freeze({ ...axis, ...(axisTitle === undefined ? {} : { title: axisTitle }), ...(numberFormat === undefined ? {} : { numberFormat }) });
  }
  const common = { ...anchor, type,
    ...(language === undefined ? {} : { language }),
    ...(title === undefined ? {} : { title }), ...(description === undefined ? {} : { description }),
    ...(input.legend === undefined ? {} : { legend: input.legend }), ...(dataLabels === undefined ? {} : { dataLabels }),
    ...(categoryAxis === undefined ? {} : { categoryAxis }), ...(valueAxis === undefined ? {} : { valueAxis }) };
  if (bound) {
    if (!resolveBinding) throw new TypeError("A chart range must belong to the current workbook.");
    const { descriptor, retainText } = resolveBinding(bound);
    const binding = prepareChartBinding(bound, descriptor, { availablePoints, availableText: availableText - textCharacters, retainText, policy }, signal, type);
    return Object.freeze({ ...common, data: binding.data, binding, points, textCharacters: textCharacters + binding.textCharacters });
  }
  // Length budgets precede allocation; each source property/array element is read once.
  const categoryValues = categories as readonly unknown[];
  const copiedCategories: string[] = [], copiedSeries = [];
  for (let i = 0; i < categoryCount; i++) { if ((i & 255) === 0) checkAbort(signal); copiedCategories.push(text(categoryValues[i], "Chart category")); }
  for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex++) {
    const item = series[seriesIndex];
    const captured = snapshot(item, ["name", "values", "color"], "chart.series");
    const name = text(captured.name, "Chart series name"), values = captured.values, color = captured.color;
    if (!Array.isArray(values) || values.length !== categoryCount) throw new TypeError("Chart series values must match the category count.");
    if (color !== undefined && (typeof color !== "string" || !/^#?[0-9a-fA-F]{6}$/.test(color))) throw new TypeError("Chart colors must be opaque six-digit RGB hex values.");
    let positive = false;
    const copiedValues: number[] = [];
    for (let i = 0; i < categoryCount; i++) {
      if ((i & 255) === 0) checkAbort(signal);
      const value: unknown = values[i];
      if (typeof value !== "number" || !Number.isFinite(value)) throw new TypeError("Chart values must be finite numbers.");
      if (type === "pie" && value < 0) throw new RangeError("Pie charts require nonnegative values.");
      if (value > 0) positive = true;
      copiedValues.push(value);
    }
    if (type === "pie" && !positive) throw new RangeError("A pie chart needs at least one positive value.");
    copiedSeries.push(Object.freeze({ name, values: Object.freeze(copiedValues), ...(color === undefined ? {} : { color: color.replace(/^#/, "").toUpperCase() }) }));
  }
  return Object.freeze({ ...common,
    data: Object.freeze({ categories: Object.freeze(copiedCategories), series: Object.freeze(copiedSeries) }),
    points, textCharacters });
}

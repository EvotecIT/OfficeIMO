import { NotSupportedError } from "../../core/errors.js";
import type { ChartForgeXCapture } from "./types.js";

export type ChartXKind = "date" | "number" | "category";
export interface PreparedPoint {
  readonly x: Date | number | string | null;
  /** Identity of the x value for wide layouts; null when absent. */
  readonly key: number | string | null;
  readonly xCaption: string | undefined;
  readonly y: number | null;
  readonly caption: string | undefined;
}
export interface PreparedSeries {
  readonly id: string;
  readonly name: string;
  readonly unit: string | undefined;
  readonly format: string | undefined;
  readonly points: readonly PreparedPoint[];
}
export interface PreparedChart {
  readonly title: string | undefined;
  readonly description: string | undefined;
  readonly kind: ChartXKind;
  readonly xTitle: string;
  readonly xFormat: string | undefined;
  readonly yTitle: string;
  readonly yUnit: string | undefined;
  readonly yFormat: string | undefined;
  readonly series: readonly PreparedSeries[];
  readonly svg: string | undefined;
}

function record(value: unknown, name: string): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError("Chart capture " + name + " must be an object.");
  return value as Record<string, unknown>;
}
function text(value: unknown, name: string, required = false): string | undefined {
  if (value === undefined || value === null) { if (required) throw new TypeError("Chart capture " + name + " is required."); return undefined; }
  if (typeof value !== "string") throw new TypeError("Chart capture " + name + " must be a string.");
  return value;
}
const isoDate = /^\d{4}-\d{2}-\d{2}(?:[T ]\d{2}:\d{2}(?::\d{2}(?:\.\d{1,9})?)?)?(?:Z|[+-]\d{2}:?\d{2})?$/;
/** Dates without an offset are UTC wall-clock values, the way ChartForgeX serializes time axes, never the reader's local time. */
function isoInstant(value: string): string {
  const normalized = value.replace(" ", "T").replace(/([+-]\d{2})(\d{2})$/, "$1:$2").replace(/(\.\d{3})\d+/, "$1");
  return normalized.length === 10 || /(Z|[+-]\d{2}:\d{2})$/.test(normalized) ? normalized : normalized + "Z";
}

/** Read the capture once into owned records, so later caller mutation cannot change an export in progress. */
export function prepareChart(input: ChartForgeXCapture, maxPoints = 1048575): PreparedChart {
  const capture = record(input, "");
  if (capture.version !== 1) throw new NotSupportedError("Only ChartForgeX capture version 1 is supported.");
  const axis = record(capture.x, "x"), kind = axis.kind;
  if (kind !== "date" && kind !== "number" && kind !== "category") throw new TypeError("Chart capture x.kind must be date, number or category.");
  const yAxis = capture.y === undefined ? {} : record(capture.y, "y");
  if (!Array.isArray(capture.series)) throw new TypeError("Chart capture series must be an array.");
  let count = 0;
  const series = Array.from(capture.series as unknown[], (entry, index): PreparedSeries => {
    const value = record(entry, "series[" + index + "]"), points = value.points;
    if (!Array.isArray(points)) throw new TypeError("Chart capture series[" + index + "].points must be an array.");
    if ((count += points.length) > maxPoints) throw new RangeError("Chart capture has more than " + maxPoints + " points.");
    return {
      id: text(value.id, "series[" + index + "].id", true)!, name: text(value.name, "series[" + index + "].name", true)!,
      unit: text(value.unit, "series unit"), format: text(value.format, "series format"),
      points: Array.from(points as unknown[], (raw, point): PreparedPoint => {
        const item = record(raw, "point"), where = "series[" + index + "].points[" + point + "]";
        const y = item.y;
        if (y !== null && y !== undefined && (typeof y !== "number" || !Number.isFinite(y))) throw new TypeError(where + ".y must be a finite number or null.");
        let x: Date | number | string | null = null, key: number | string | null = null;
        const source = item.x;
        if (source !== null && source !== undefined) {
          if (kind === "date") {
            let time = NaN;
            if (typeof source === "string" && isoDate.test(source.trim())) time = Date.parse(isoInstant(source.trim()));
            if (typeof source === "object") {
              try { time = Date.prototype.getTime.call(source); } catch { /* Not a Date. */ }
            }
            if (!Number.isFinite(time)) throw new TypeError(where + ".x must be an ISO 8601 date string or a valid Date.");
            x = new Date(time); key = time;
          } else if (kind === "number") {
            if (typeof source !== "number" || !Number.isFinite(source)) throw new TypeError(where + ".x must be a finite number.");
            x = key = source;
          } else {
            if (typeof source !== "string" && typeof source !== "number") throw new TypeError(where + ".x must be a category label.");
            x = key = String(source);
          }
        }
        return { x, key, xCaption: text(item.xCaption, where + ".xCaption"), y: y ?? null, caption: text(item.caption, where + ".caption") };
      })
    };
  });
  const svg = capture.svg;
  const markup = svg && typeof svg === "object" ? (svg as { outerHTML?: unknown }).outerHTML : svg;
  if (svg != null && typeof markup !== "string") throw new TypeError("Chart capture svg must be markup or an element.");
  return {
    title: text(capture.title, "title"), description: text(capture.description, "description"), kind,
    xTitle: text(axis.title, "x.title") || "X", xFormat: text(axis.format, "x.format"),
    yTitle: text(yAxis.title, "y.title") || "Value", yUnit: text(yAxis.unit, "y.unit"), yFormat: text(yAxis.format, "y.format"),
    series, svg: typeof markup === "string" ? markup : undefined
  };
}

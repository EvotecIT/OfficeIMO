import type { ExportLimits, ExportProgress } from "../../core/index.js";
import type { XlsxOptions } from "../../xlsx/index.js";
import type { PdfDocumentOptions } from "../../pdf/index.js";
import type { PdfDrawingFont } from "../../pdf/drawing-types.js";

/** One observation. `x` matches the capture's x kind: an ISO 8601 string or Date, a finite number, or a category label. */
export interface ChartForgeXPoint {
  readonly x: string | number | Date | null;
  /** The x value as the chart displays it, for example `Mar 2026`. */
  readonly xCaption?: string;
  readonly y: number | null;
  /** The value as the chart displays it, for example `1.2k ms`. */
  readonly caption?: string;
}
export interface ChartForgeXSeries {
  /** Stable series identity; names are display text and may repeat. */
  readonly id: string;
  readonly name: string;
  /** Unit appended to the value heading, for example `ms`. Overrides the y axis unit. */
  readonly unit?: string;
  /** Native Excel number format for this series' values. Overrides the y axis format. */
  readonly format?: string;
  readonly points: readonly ChartForgeXPoint[];
}
export interface ChartForgeXAxis {
  readonly title?: string;
  readonly unit?: string;
  /** Native Excel number format, for example `0.0` or `yyyy-mm-dd`. */
  readonly format?: string;
}
/** A plain chart capture, as ChartForgeX's browser runtime produces it. OfficeIMO never reads the live chart. */
export interface ChartForgeXCapture {
  readonly version: 1;
  readonly id?: string;
  readonly title?: string;
  readonly description?: string;
  readonly x: ChartForgeXAxis & { readonly kind: "date" | "number" | "category" };
  readonly y?: ChartForgeXAxis;
  readonly series: readonly ChartForgeXSeries[];
  /** Serialized SVG whose paint is resolved inline (style declarations or presentation attributes). Required for PDF. */
  readonly svg?: string | { readonly outerHTML: string };
}
export type ChartForgeXFormat = "xlsx" | "pdf";
export interface ChartForgeXDiagnostic {
  /** For example `OPACITY_FLATTENED`, `GRADIENT_FLATTENED`, `CLIP_IGNORED`, `IMAGE_SKIPPED` or `TEXT_SUBSTITUTED`. */
  readonly code: string;
  /** The SVG element name the diagnostic applies to. */
  readonly element?: string;
}
export interface ChartForgeXSvgOptions {
  /** SVG features without a vector equivalent are skipped with a diagnostic by default; `reject` fails instead. */
  readonly unsupported?: "skip" | "reject";
  /** Drawing font slots the PDF can paint. Text in another slot falls back to `regular`. Default: all four. */
  readonly fonts?: readonly PdfDrawingFont[];
  /** Map text to the standard fonts' WinAnsi repertoire, substituting other characters. Default true. */
  readonly winAnsi?: boolean;
  readonly onDiagnostic?: (diagnostic: ChartForgeXDiagnostic) => void;
}
export interface ChartForgeXWriteOptions {
  /** `wide`: one row per x value and a column per series. `long`: one row per point. `auto` (default) picks wide unless
   * a series repeats an x value or a date/number point has no x. */
  readonly layout?: "auto" | "wide" | "long";
  readonly signal?: AbortSignal;
  readonly limits?: ExportLimits;
  readonly onProgress?: (progress: ExportProgress) => void;
  readonly onDiagnostic?: (diagnostic: ChartForgeXDiagnostic) => void;
  readonly unsupported?: "skip" | "reject";
  readonly xlsx?: Pick<XlsxOptions, "dateMode" | "compression" | "creator" | "created" | "modified"> & { readonly sheetName?: string };
  readonly pdf?: Pick<PdfDocumentOptions, "pageSize" | "orientation" | "margins" | "fonts" | "tagged" | "compression" | "pageNumbers" |
    "pageHeader" | "pageFooter" | "fontSize" | "limits">;
}

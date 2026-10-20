import type { WorksheetAnchor } from "./attachments.js";

/** Native, editable Excel chart families. Names follow OfficeIMO's shared chart vocabulary. */
export type ChartType = "columnClustered" | "barClustered" | "line" | "pie";

/** Values align with the categories. Strings, missing points and nonfinite values are rejected. */
export interface ChartSeries {
  readonly name: string;
  readonly values: readonly number[];
  /** Opaque RGB color, with or without a leading #. Excel supplies its palette when omitted. */
  readonly color?: string;
}

/** An independent snapshot; the workbook writes its editable source cells on a hidden data sheet. */
export interface ChartData {
  readonly categories: readonly string[];
  readonly series: readonly ChartSeries[];
}

/** Labels are native Excel labels and remain editable after export. */
export interface ChartDataLabels {
  readonly showValue?: boolean;
  readonly showCategoryName?: boolean;
  readonly showSeriesName?: boolean;
  /** Percent labels are supported only by pie charts. */
  readonly showPercent?: boolean;
  readonly position?: "center" | "insideEnd" | "outsideEnd" | "bestFit" | "above" | "below" | "left" | "right";
  readonly numberFormat?: string;
}

export interface ChartCategoryAxis { readonly title?: string; }
export interface ChartValueAxis {
  readonly title?: string;
  readonly minimum?: number;
  readonly maximum?: number;
  readonly numberFormat?: string;
}

/** A native chart placed in worksheet coordinates, separate from PNG or portable drawing rendering. */
export interface WorksheetChart extends WorksheetAnchor {
  readonly type: ChartType;
  readonly data: ChartData;
  readonly title?: string;
  /** The legend is on the right by default. */
  readonly legend?: "right" | "bottom" | "left" | "top" | false;
  readonly dataLabels?: ChartDataLabels;
  /** Pie charts have no axes and reject axis options. */
  readonly categoryAxis?: ChartCategoryAxis;
  readonly valueAxis?: ChartValueAxis;
}

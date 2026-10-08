import type { Column, ExportValue, CellPresentation, CellValue, ColumnValueContext, StreamOptions, ExportLimits, TableSpanRows } from "../core/index.js";
import type { PdfFont } from "./font.js";

/** PDF lengths are points (72 points per inch). */
export interface PdfPageSize { readonly width: number; readonly height: number; }
export interface PdfMargins { readonly top: number; readonly right: number; readonly bottom: number; readonly left: number; }
export interface PdfFonts { readonly regular: PdfFont; readonly bold?: PdfFont; readonly italic?: PdfFont; readonly boldItalic?: PdfFont; }
export interface PdfLimits extends ExportLimits {
  readonly maxPages?: number;
  readonly maxColumns?: number;
  /** Bounds one projected cell before wrapping; text is never truncated. */
  readonly maxCellCharacters?: number;
  /** Bounds one row's wrapped layout, including a row split over many pages. */
  readonly maxRowLines?: number;
  readonly maxFontBytes?: number;
  readonly maxPageBytes?: number;
  /** Actual link annotations, including repeated headings and continued cell fragments. */
  readonly maxHyperlinks?: number;
}
export type PdfTotal = "sum" | "count" | "average" | "min" | "max";
export interface PdfTableFooter {
  readonly values?: readonly ExportValue[];
  /** Keys are declared column keys, or unique headers when a key is absent. */
  readonly totals?: Readonly<Record<string, PdfTotal>>;
  /** Structured alternative to values/totals; includes horizontal and vertical spans. */
  readonly rows?: TableSpanRows;
}
export interface PdfPageContext { readonly pageNumber: number; }
export interface PdfOptions<T = never> extends Omit<StreamOptions, "limits"> {
  readonly columns: readonly Column<T>[];
  readonly limits?: PdfLimits;
  readonly pageSize?: "A3" | "A4" | "A5" | "LETTER" | "LEGAL" | "TABLOID" | PdfPageSize;
  readonly orientation?: "portrait" | "landscape";
  readonly margins?: number | Partial<PdfMargins>;
  /** Standard Helvetica is used when omitted. Supply TrueType fonts for Unicode text. */
  readonly fonts?: PdfFonts;
  readonly fontSize?: number;
  /** First-page title. Empty text is omitted without reserving space or consuming a cell. */
  readonly title?: string;
  /** Text before the first table heading. Empty text is omitted. */
  readonly messageTop?: string;
  /** Text after the table. Empty text is omitted. */
  readonly messageBottom?: string;
  readonly includeHeader?: boolean;
  /** Overrides Column.groups/headers with a small explicit span matrix, repeated on each page. */
  readonly headerRows?: TableSpanRows;
  readonly footer?: PdfTableFooter;
  readonly headerPresentation?: CellPresentation;
  readonly footerPresentation?: CellPresentation;
  readonly alternateRowColor?: string;
  readonly padding?: number;
  /** Explicit point widths. Otherwise shared Column.width is measured in zero-character widths. */
  readonly columnWidths?: readonly number[];
  /** Fit preserves every column at the requested font size by wrapping. Reject requires widths to fit. */
  readonly wideTable?: "fit" | "reject";
  readonly pageHeader?: string | ((context: PdfPageContext) => string);
  readonly pageFooter?: string | ((context: PdfPageContext) => string);
  readonly pageNumbers?: boolean;
  /** Prefer ExportCell.text for presentation, then this formatter, then ISO/scalar text. */
  readonly formatValue?: (value: CellValue, context: ColumnValueContext) => string;
  /** Native deflate when available; the fallback writes valid uncompressed PDF streams. */
  readonly compression?: boolean;
}

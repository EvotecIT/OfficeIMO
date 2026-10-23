import type { Column, ExportValue, CellPresentation, CellValue, ColumnValueContext, StreamOptions, ExportLimits, TableSpanRows } from "../core/index.js";
import type { PdfFont } from "./font.js";
import type { PdfTextRun } from "./inline-types.js";
import type { PdfTextDirection, PdfTextShaper } from "./text-shaping-types.js";
import type { BooleanText } from "../core/value-text.js";

/** Readable horizontal page bands. Every band in a window contains the same logical row fragments. */
export interface PdfTableBands {
  readonly mode: "bands";
  /** Minimum point width when columnWidths is absent. Default 72; shared Column.width can increase it. */
  readonly minimumColumnWidth?: number;
  /** Zero-based positions or exact stable column keys, repeated in selected source order on every band.
   * Numeric-only lists retain their sorted unique positional contract. Keys resolve only within the selected columns;
   * duplicate positions, including aliases, reject. Grouped tables include their leading label at numeric position zero.
   * Repeated cell contents must fit below the headings/context on one fresh page. */
  readonly repeatColumns?: readonly (number | string)[];
}
/** PDF lengths are points (72 points per inch). */
export interface PdfPageSize { readonly width: number; readonly height: number; }
export interface PdfMargins { readonly top: number; readonly right: number; readonly bottom: number; readonly left: number; }
export interface PdfFonts { readonly regular: PdfFont; readonly bold?: PdfFont; readonly italic?: PdfFont; readonly boldItalic?: PdfFont; }
/** Font family and emphasis selection for one text run; omitted fonts use standard Helvetica. */
export interface PdfTextPreflightOptions { readonly fonts?: PdfFonts; readonly bold?: boolean; readonly italic?: boolean; }
/** Opt-in logical structure for the supported report profile; does not assert PDF/UA conformance. */
export interface PdfTaggedOptions { readonly language: string; }
/** Stable reason that the current scalar-layout writer cannot render a text run. */
export type PdfTextIssueCode = "unpaired-surrogate" | "unsupported-control" | "unsupported-layout" | "embedded-font-required" | "missing-glyph";
/** Text support alone does not establish layout, resource-budget or accessibility success. */
export type PdfTextPreflightResult = { readonly supported: true } | {
  readonly supported: false;
  readonly code: PdfTextIssueCode;
  /** Zero-based UTF-16 index in the original input, before line-break and tab normalization. */
  readonly index: number;
  /** Original Unicode scalar; absent when the issue is an unpaired surrogate. */
  readonly scalar?: number;
  readonly message: string;
};
export interface PdfLimits extends ExportLimits {
  readonly maxPages?: number;
  readonly maxColumns?: number;
  /** Horizontal bands in one table; includes repeated identifier columns. Default 64. */
  readonly maxTableBands?: number;
  /** Retained logical row fragments in a band window. Smaller values flush earlier. Default 256. */
  readonly maxTableWindowRows?: number;
  /** Bounds one projected cell before wrapping; text is never truncated. */
  readonly maxCellCharacters?: number;
  /** Captured styled runs per paragraph/cell, and in a kept paragraph group. */
  readonly maxTextRuns?: number;
  /** Paragraph count in a group or kept chain. */
  readonly maxGroupBlocks?: number;
  /** Bounds one row's wrapped layout, including a row split over many pages. */
  readonly maxRowLines?: number;
  readonly maxFontBytes?: number;
  readonly maxPageBytes?: number;
  /** Actual link annotations, including repeated headings and continued cell fragments. */
  readonly maxHyperlinks?: number;
  /** Total source bytes of distinct image objects retained by a composed document. */
  readonly maxImageBytes?: number;
  /** Per-image decoded pixel ceiling, checked before allocating RGB or alpha planes. */
  readonly maxImagePixels?: number;
  /** Optional ceiling for owned image input/decoded/scanline buffers; excludes native codec memory and compressed output. */
  readonly maxImageWorkingBytes?: number;
  /** Total path commands across a composed document, checked before serializing a block. */
  readonly maxVectorCommands?: number;
  /** Total path, text and group elements across portable drawing blocks. */
  readonly maxDrawingElements?: number;
  /** Total logical structure elements created across a tagged document, including Document, rows and cells. */
  readonly maxStructureElements?: number;
  /** Total marked-content references created across the document, including continued cells. */
  readonly maxMarkedContentReferences?: number;
  /** Section boundaries across one composed document. */
  readonly maxSections?: number;
  /** Native outline entries across one composed document. */
  readonly maxBookmarks?: number;
  /** Streamed list items across one composed document. */
  readonly maxListItems?: number;
  /** Captured named paragraph styles in a composed document. */
  readonly maxParagraphStyles?: number;
}
export type PdfTotal = "sum" | "count" | "average" | "min" | "max";
export interface PdfTableFooter {
  readonly values?: readonly ExportValue[];
  /** Keys are declared column keys, or unique headers when a key is absent. */
  readonly totals?: Readonly<Record<string, PdfTotal>>;
  /** Structured alternative to values/totals; includes horizontal and vertical spans. */
  readonly rows?: TableSpanRows;
}
export interface PdfPageContext {
  readonly pageNumber: number;
  /** One-based section number; an unsectioned document uses section one. */
  readonly sectionNumber: number;
  /** One-based physical page within the current section. */
  readonly sectionPageNumber: number;
  readonly sectionTitle?: string;
}
/** Literal text around the current/total page counts; no locale is inferred. */
export interface PdfPageNumberLabels {
  /** Text before the current page number. Default "Page ". Empty text is valid. */
  readonly prefix?: string;
  /** Text between the current and total page numbers. Default " of ". */
  readonly separator?: string;
  /** Text after the total page count. Default empty. */
  readonly suffix?: string;
}
export interface PdfOptions<T = never> extends Omit<StreamOptions, "limits"> {
  /** Optional provider for composed paragraphs and table cells. Legacy decorations remain scalar text. */
  readonly textShaper?: PdfTextShaper;
  /** Base direction of each shaped physical paragraph; defaults to auto. */
  readonly textDirection?: PdfTextDirection;
  readonly columns: readonly Column<T>[];
  readonly limits?: PdfLimits;
  readonly pageSize?: "A3" | "A4" | "A5" | "LETTER" | "LEGAL" | "TABLOID" | PdfPageSize;
  readonly orientation?: "portrait" | "landscape";
  readonly margins?: number | Partial<PdfMargins>;
  /** Standard Helvetica is used when omitted. Supply TrueType fonts for Unicode text. */
  readonly fonts?: PdfFonts;
  /** Logical reading order, simple table semantics and explicit figure alternatives. */
  readonly tagged?: PdfTaggedOptions;
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
  readonly wideTable?: "fit" | "reject" | PdfTableBands;
  readonly pageHeader?: string | ((context: PdfPageContext) => string);
  readonly pageFooter?: string | ((context: PdfPageContext) => string);
  readonly pageNumbers?: boolean;
  /** Captured page-number literals. English defaults preserve the existing display contract. */
  readonly pageNumberLabels?: PdfPageNumberLabels;
  /** Explicit boolean display labels. Defaults are the existing English True/False text. */
  readonly booleanText?: BooleanText;
  /** Prefer ExportCell.text for presentation, then this formatter, then ISO/scalar text. */
  readonly formatValue?: (value: CellValue, context: ColumnValueContext) => string;
  /** Styled display runs when ExportCell.text is absent; must return synchronously. */
  readonly formatRuns?: (value: CellValue, context: ColumnValueContext) => readonly PdfTextRun[];
  /** Native deflate when available; the fallback writes valid uncompressed PDF streams. */
  readonly compression?: boolean;
}

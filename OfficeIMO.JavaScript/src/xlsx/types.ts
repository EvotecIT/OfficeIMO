import type { Column, CellValue, StreamOptions } from "../core/index.js";
import type { ExportLimits } from "../core/limits.js";
import type { InvalidCharacterPolicy } from "../xml/index.js";
import type { Compression } from "../zip/index.js";
import type { PackagePart, Relationship } from "../opc/index.js";
import type { CoreProperties, AppProperties } from "../opc/index.js";
import type { Cell } from "./values.js";
import type { CellStyle } from "./styles.js";
import type { Hyperlink } from "./attachments.js";
import type { ExportValue } from "../core/presentation.js";

export type XlsxRow = readonly (ExportValue | Cell)[] | Readonly<Record<string, ExportValue | Cell>>;
export type XlsxRows = Iterable<XlsxRow> | AsyncIterable<XlsxRow>;
export interface CellWriterContext { readonly column: Column; readonly row: number; readonly columnIndex: number; readonly sheetName: string; }
/** Convert domain column values to plain values or styled Cells; raw XML is never accepted. */
export type CellValueWriter = (value: CellValue, context: CellWriterContext) => ExportValue | Cell;
/** One-based worksheet row number, including the header; values follow the declared column order. */
export interface RowStyleContext { readonly row: number; readonly values: readonly CellValue[]; readonly sheetName: string; }
export interface CellStyleContext extends RowStyleContext { readonly value: CellValue; readonly column: Column; readonly columnIndex: number; }
/** A real Excel table over this worksheet's header and data rows. Empty exports retain headers without a table part. */
export interface TableOptions {
  readonly name?: string;
  /** Built-in Excel TableStyleLight1..21, TableStyleMedium1..28 or TableStyleDark1..11. */
  readonly style?: string;
  readonly bandedRows?: boolean;
  readonly bandedColumns?: boolean;
  readonly firstColumn?: boolean;
  readonly lastColumn?: boolean;
}
export interface WorkbookOptions extends StreamOptions, CoreProperties {
  /** Selected before rows are written. Streamed worksheets are completed in order. The caller owns the sink. */
  readonly sink?: import("../core/sinks.js").ByteSink;
  readonly limits?: WorkbookLimits;
  readonly dateMode?: "local" | "utc";
  readonly compression?: Compression;
  readonly invalidCharacterPolicy?: InvalidCharacterPolicy;
  readonly appProperties?: AppProperties;
  readonly cellValueWriters?: Readonly<Record<string, CellValueWriter>>;
  /** Reject by default. Preserve keeps a bounded preview and the full text in numbered overflow-sheet chunks. */
  readonly oversizedText?: "reject" | "preserve";
}
export interface WorkbookLimits extends ExportLimits {
  readonly maxStyles?: number;
  readonly maxSheets?: number;
  /** Retained UTF-16 units before the overflow sheet is emitted; defaults to 4,000,000. */
  readonly maxOverflowCharacters?: number;
  readonly maxHyperlinks?: number;
  readonly maxImageBytes?: number;
  readonly maxBufferedCells?: number;
  readonly maxBufferedCharacters?: number;
}
export interface XlsxExportResult { readonly rows: number; readonly sheets: number; readonly bytes: number; }
export interface SheetOptions {
  readonly columns?: readonly Column[];
  readonly includeHeader?: boolean;
  readonly freezeHeader?: boolean;
  readonly autoFilter?: boolean;
  readonly boldHeader?: boolean;
  readonly headerFill?: string;
  readonly headerStyle?: number;
  /** Presentation patches retain column number/date formats. The second data row gets this patch first. */
  readonly alternatingRowStyle?: CellStyle;
  readonly rowStyle?: (context: RowStyleContext) => CellStyle | undefined;
  readonly cellStyle?: (context: CellStyleContext) => CellStyle | undefined;
  readonly rowHeight?: number;
  readonly headerHeight?: number;
  /** Fallback width in characters. Native tables default to 20 when no column width is supplied. */
  readonly defaultColumnWidth?: number;
  readonly freezeColumns?: number;
  readonly table?: TableOptions;
  readonly autoSize?: AutoSizeOptions;
  readonly footer?: FooterOptions;
  readonly print?: PrintOptions;
  /** Reserved and rejected until implemented. Presence, including an empty array, throws. */
  readonly mergedCells?: readonly unknown[];
  readonly hyperlinks?: readonly Hyperlink[];
  readonly conditionalFormats?: readonly unknown[];
  readonly dataValidation?: readonly unknown[];
}
/** Widths use a bounded leading sample and an approximate character count, not font measurement. */
export interface AutoSizeOptions { readonly sampleRows?: number; readonly minWidth?: number; readonly maxWidth?: number; }
export type TotalOperation = "sum" | "count" | "average" | "min" | "max";
export interface FooterOptions {
  /** One explicit value per declared column; computed totals replace the corresponding value. */
  readonly values?: readonly (ExportValue | Cell)[];
  readonly totals?: Readonly<Record<string, TotalOperation>>;
  readonly style?: import("./styles.js").CellStyle;
}
export interface PrintOptions {
  readonly paper?: "A4" | "Letter";
  readonly orientation?: "portrait" | "landscape";
  readonly fitToWidth?: number;
  readonly fitToHeight?: number;
  readonly repeatHeaders?: boolean;
  readonly margins?: Readonly<Partial<Record<"left" | "right" | "top" | "bottom" | "header" | "footer", number>>>;
  /** Literal center header/footer text. Ampersands are escaped, so Excel control codes are not interpreted. */
  readonly header?: string;
  readonly footer?: string;
}
export interface ExtraPart extends PackagePart {
  /** Optional relationship to this part, from the workbook unless a source is supplied. */
  readonly relationship?: Omit<Relationship, "target" | "external"> & { readonly source?: string };
}

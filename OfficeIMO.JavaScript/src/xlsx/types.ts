import type { Column, CellValue, StreamOptions } from "../core/index.js";
import type { InvalidCharacterPolicy } from "../xml/index.js";
import type { Compression } from "../zip/index.js";
import type { PackagePart, Relationship } from "../opc/index.js";
import type { CoreProperties, AppProperties } from "../opc/index.js";
import type { Cell } from "./values.js";

export type XlsxRow = readonly (CellValue | Cell)[] | Readonly<Record<string, CellValue | Cell>>;
export type XlsxRows = Iterable<XlsxRow> | AsyncIterable<XlsxRow>;
export interface CellWriterContext { readonly column: Column; readonly row: number; readonly columnIndex: number; readonly sheetName: string; }
/** Convert domain column values to plain values or styled Cells; raw XML is never accepted. */
export type CellValueWriter = (value: CellValue, context: CellWriterContext) => CellValue | Cell;
export interface WorkbookOptions extends StreamOptions, CoreProperties {
  readonly dateMode?: "local" | "utc";
  readonly compression?: Compression;
  readonly invalidCharacterPolicy?: InvalidCharacterPolicy;
  readonly appProperties?: AppProperties;
  readonly cellValueWriters?: Readonly<Record<string, CellValueWriter>>;
}
export interface SheetOptions {
  readonly columns?: readonly Column[];
  readonly includeHeader?: boolean;
  readonly freezeHeader?: boolean;
  readonly autoFilter?: boolean;
  readonly boldHeader?: boolean;
  readonly headerFill?: string;
  /** Reserved and rejected until implemented. Presence, including an empty array, throws. */
  readonly mergedCells?: readonly unknown[];
  readonly hyperlinks?: readonly unknown[];
  readonly conditionalFormats?: readonly unknown[];
  readonly dataValidation?: readonly unknown[];
}
export interface ExtraPart extends PackagePart {
  /** Optional relationship to this part, from the workbook unless a source is supplied. */
  readonly relationship?: Omit<Relationship, "target" | "external"> & { readonly source?: string };
}

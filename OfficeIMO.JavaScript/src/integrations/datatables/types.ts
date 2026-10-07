import type { Column, ExportValue, StreamOptions, TableSpanRows } from "../../core/index.js";
import type { CsvOptions } from "../../csv/index.js";
import type { PortableSheetOptions, PortableWorkbookOptions, WorkbookLimits } from "../../xlsx/index.js";
import type { PdfOptions, PdfLimits } from "../../pdf/index.js";

/** Structural interoperability boundary; importing this adapter never imports DataTables. */
export type DataTablesMethod = (...args: never[]) => unknown;
export interface DataTablesApi {
  readonly rows: DataTablesMethod;
  readonly columns: DataTablesMethod;
  readonly cells: DataTablesMethod;
  readonly page: { readonly info: DataTablesMethod };
  readonly buttons: { readonly exportData: DataTablesMethod; readonly exportInfo?: DataTablesMethod };
}
export interface DataTablesHost {
  /** Checked at runtime; older Buttons declarations omit the public stripData helper. */
  readonly Buttons: object;
  readonly ext: { readonly buttons: Record<string, unknown> };
}
export interface DataTablesCellContext {
  /** DataTables indexes, after any column reordering. */
  readonly sourceRowIndex: number;
  readonly sourceColumnIndex: number;
  /** Zero-based position in the selected export. */
  readonly rowIndex: number;
  readonly columnIndex: number;
}
export interface DataTablesFormat {
  readonly header?: (value: unknown, column: number, node: unknown) => unknown;
  readonly footer?: (value: unknown, column: number, node: unknown) => unknown;
  readonly body?: (value: unknown, row: number, column: number, node: unknown) => unknown;
}
/** Passed to the installed Buttons implementation in compatibility mode. */
export interface DataTablesExportOptions {
  readonly rows?: unknown;
  readonly columns?: unknown;
  readonly modifier?: Readonly<Record<string, unknown>>;
  readonly orthogonal?: string;
  readonly stripHtml?: boolean;
  readonly stripNewlines?: boolean;
  readonly decodeEntities?: boolean;
  readonly escapeExcelFormula?: boolean;
  readonly trim?: boolean;
  readonly format?: DataTablesFormat;
  /** Whole-matrix customization is supported only by compatibility mode. */
  readonly customizeData?: (data: unknown) => void;
}
export interface DataTablesOptions extends StreamOptions {
  /** Batched avoids Buttons' complete body matrix; compatibility delegates gathering to Buttons. */
  readonly mode?: "batched" | "compatibility";
  readonly exportOptions?: DataTablesExportOptions;
  /** Maximum rows per projection batch, default 1,024, at most 4,096. */
  readonly batchRows?: number;
  /** Maximum cells per projection batch, default 65,536. One row must fit. */
  readonly maxBatchCells?: number;
  /** Grouped accepts horizontal spans; leaf selects one row; structured preserves rectangles for PDF. */
  readonly headings?: "grouped" | "leaf" | "structured";
  readonly includeFooter?: boolean;
  /** Server-side tables require explicit acknowledgement that only loaded rows are available. */
  readonly serverSide?: "reject" | "loaded";
  /** Overrides keyed by DataTables column index, independent of the selected export position. */
  readonly columnOptions?: Readonly<Record<number, Partial<Omit<Column, "groups" | "value">> & { readonly style?: never }>>;
  /** Resolve portable values/presentation once. Results must be synchronous scalar values or ExportCells. */
  readonly project?: (value: ExportValue, context: DataTablesCellContext) => ExportValue;
}
/** Selection and heading metadata are captured immediately; batched body values are read during iteration. */
export interface DataTablesExport {
  readonly columns: readonly Column<readonly ExportValue[]>[];
  readonly headers: readonly (readonly ExportValue[])[];
  readonly footer: readonly ExportValue[] | undefined;
  readonly headerStructure?: TableSpanRows;
  readonly footerStructure?: TableSpanRows;
  readonly rowCount: number;
  /** Single-use source. Keep the table's data stable until iteration finishes. */
  readonly rows: AsyncIterable<readonly ExportValue[]>;
}
export interface DataTablesWriteOptions extends DataTablesOptions {
  readonly limits?: WorkbookLimits & PdfLimits;
  readonly sheetName?: string;
  readonly workbook?: Omit<PortableWorkbookOptions, "signal" | "onProgress" | "limits">;
  readonly sheet?: Omit<PortableSheetOptions, "includeHeader">;
  readonly csv?: Omit<CsvOptions, "columns" | "includeHeader" | "signal" | "onProgress" | "limits">;
  readonly pdf?: Omit<PdfOptions, "columns" | "signal" | "onProgress" | "limits">;
}
export interface DataTablesButtonOptions extends DataTablesWriteOptions {
  readonly filename?: string | ((configuration: unknown, table: DataTablesApi) => string);
  /** Receives failures after Buttons' completion callback has cleared its processing state. */
  readonly onError?: (error: unknown) => void | Promise<void>;
  /** Override destination delivery, for example to retain a download in an application. */
  readonly save?: (blob: Blob, filename: string) => void | Promise<void>;
}

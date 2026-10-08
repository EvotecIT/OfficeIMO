import type { CellPresentation, ColumnSettings, ExportLimits, ExportProgress, ExportValue } from "../../core/index.js";
import type { ExportLink } from "../../core/links.js";
import type { CsvOptions } from "../../csv/index.js";
import type { XlsxOptions } from "../../xlsx/index.js";
import type { PdfOptions } from "../../pdf/index.js";

/** Structural portable capture. CanopyX owns queries, selection, presenters and revision validation. */
export interface CanopyExportCell {
  readonly value: string | number | boolean | null;
  readonly text: string;
  readonly tone?: string;
  readonly link?: ExportLink;
  readonly diagnostics?: readonly string[];
}
export interface CanopyExportColumn {
  readonly id: string;
  readonly title: string;
  readonly kind: "text" | "number" | "datetime";
  readonly width?: { readonly unit: "css-px"; readonly preferred: number; readonly minimum: number; readonly maximum?: number };
  readonly wrap?: boolean;
  readonly alignment?: "left" | "right";
  readonly diagnostics?: readonly string[];
}
export interface CanopyExportRow {
  readonly id: string;
  readonly cells: Readonly<Record<string, CanopyExportCell>>;
  readonly tone?: string;
  readonly diagnostics?: readonly string[];
}
export interface CanopyCapture {
  readonly request: {
    readonly columns: readonly CanopyExportColumn[];
    readonly recordCount: number | null;
    readonly values: "raw" | "display";
    readonly presentation: "text" | "semantic";
    readonly timeZone: "utc" | "local";
    readonly revision: string;
  };
  /** Fresh serial enumeration. The signal must reach the source's in-flight page requests. */
  rows(options?: { readonly signal?: AbortSignal }): AsyncIterable<CanopyExportRow>;
}
export type CanopyFormat = "xlsx" | "csv" | "pdf";
export interface CanopyDiagnostic {
  readonly code: string;
  readonly rowId?: string;
  readonly columnId?: string;
}
export interface CanopySourceOptions {
  /** Map semantic tones explicitly; CSS and DOM rules are never inferred. Cell tone overrides row tone. */
  readonly tones?: Readonly<Record<string, CellPresentation>>;
  /** Selected-column overrides. Unselected IDs are ignored; widths use Excel zero-character units, not CSS pixels. */
  readonly columnOptions?: Readonly<Record<string, Omit<Partial<ColumnSettings>, "header" | "type">>>;
  /** XLSX/PDF reject unsupported presentation by default. CSV defaults to text because it has no visual styling. */
  readonly unsupportedPresentation?: "reject" | "text";
  /** XLSX: preserve uses typed Dates where exact and literal ISO text for greater precision. */
  readonly datetime?: "preserve" | "typed" | "text";
  /** The same clock used by the XLSX writer; defaults to the native capture's timeZone. */
  readonly xlsx?: Pick<XlsxOptions, "dateMode">;
  readonly signal?: AbortSignal;
  /** Synchronous notification; throwing aborts enumeration. No unbounded diagnostic list is retained. */
  readonly onDiagnostic?: (diagnostic: CanopyDiagnostic) => void;
}
/** Each call to rows opens an independent native capture enumeration. No whole-table matrix is retained. */
export interface CanopyExport {
  readonly columns: readonly (ColumnSettings & { readonly key: string })[];
  readonly rowCount: number | null;
  rows(): AsyncIterable<readonly ExportValue[]>;
}
export interface CanopyWriteOptions extends CanopySourceOptions {
  readonly limits?: ExportLimits;
  readonly onProgress?: (progress: ExportProgress) => void;
  readonly xlsx?: Omit<XlsxOptions, "columns" | "signal" | "onProgress">;
  readonly csv?: Omit<CsvOptions, "columns" | "signal" | "onProgress" | "valueMode">;
  readonly pdf?: Omit<PdfOptions, "columns" | "signal" | "onProgress">;
}

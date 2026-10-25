import type { CellPresentation, ColumnSettings, ExportLimits, ExportProgress, ExportValue } from "../../core/index.js";
import type { ExportLink } from "../../core/links.js";
import type { CsvOptions } from "../../csv/index.js";
import type { XlsxOptions } from "../../xlsx/index.js";
import type { PdfOptions, GroupedPdfOptions } from "../../pdf/index.js";

/** Structural portable capture. CanopyX owns queries, selection, presenters and revision validation. */
export interface CanopyExportCell {
  readonly value: string | number | boolean | null;
  readonly text: string;
  readonly tone?: string;
  readonly link?: ExportLink;
  readonly diagnostics?: readonly string[];
}
/** CanopyX's declarative numeric intent; this is never parsed as native Excel syntax. */
export interface CanopyNumberFormat {
  readonly minimumFractionDigits?: number;
  readonly maximumFractionDigits?: number;
  readonly prefix?: string;
  readonly suffix?: string;
}
export interface CanopyExportColumn {
  readonly id: string;
  readonly title: string;
  readonly kind: "text" | "number" | "datetime";
  readonly width?: { readonly unit: "css-px"; readonly preferred: number; readonly minimum: number; readonly maximum?: number };
  readonly pinned?: "left" | "right" | null;
  readonly format?: CanopyNumberFormat;
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
/** Captured half-open detail ranges, never ranges from a different export scope. */
export interface CanopyExportGroup {
  readonly key: string;
  readonly value: string | number | boolean | null;
  readonly text?: string;
  readonly offset: number;
  readonly count: number;
  readonly aggregates: Readonly<Record<string, number | null>>;
  readonly groups?: readonly CanopyExportGroup[];
}
export interface CanopyExportGrouping {
  readonly levels: readonly { readonly column: string; readonly direction: "asc" | "desc"; readonly title?: string }[];
  readonly aggregates: readonly { readonly id: string; readonly column: string; readonly operation: string; readonly label?: string }[];
  readonly totals: Readonly<Record<string, number | null>>;
  readonly groups: readonly CanopyExportGroup[];
}
export interface CanopyRuleFilter {
  readonly column: string;
  /** Native numeric/typed-UTC-date comparisons and null tests are bounded; unrepresentable type/culture profiles need explicit text fallback. */
  readonly operator: string;
  readonly values?: readonly (string | number | boolean)[];
}
export interface CanopyRulePredicate {
  readonly operator: string;
  readonly filters?: readonly CanopyRuleFilter[];
  readonly groups?: readonly CanopyRulePredicate[];
}
export interface CanopyFormatRule {
  readonly id: string;
  readonly enabled?: boolean;
  readonly stopIfTrue?: boolean;
  readonly column?: string;
  readonly when?: CanopyRulePredicate;
  readonly target?: "row" | "cell" | "columns";
  readonly style: {
    readonly tone?: string;
    readonly bar?: { readonly minimum?: number; readonly maximum?: number; readonly tone?: string; readonly showValue?: boolean };
    readonly scale?: { readonly stops: readonly { readonly type: string; readonly value?: number; readonly tone: string }[] };
  };
}
export interface CanopyCapture {
  readonly request: {
    readonly columns: readonly CanopyExportColumn[];
    readonly recordCount: number | null;
    readonly values: "raw" | "display";
    readonly presentation: "text" | "semantic";
    readonly timeZone: "utc" | "local";
    readonly revision: string;
    readonly scope?: "filtered" | "selected" | "all";
    readonly total?: number | null;
    readonly matched?: number | null;
    readonly selection?: readonly string[];
    /** The actual accepted export query. OfficeIMO never evaluates or substitutes a live-view query. */
    readonly query?: object;
    readonly generatedAt?: string;
    readonly grouping?: CanopyExportGrouping;
    readonly formatRules?: readonly CanopyFormatRule[];
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
  /** Map semantic tones explicitly; canonical neutral defaults to no override. Cell tone overrides row tone. */
  readonly tones?: Readonly<Record<string, CellPresentation>>;
  /** Selected-column overrides. Unselected IDs are ignored. Native format overrides remain literal Excel syntax. */
  readonly columnOptions?: Readonly<Record<string, Omit<Partial<ColumnSettings>, "header" | "type">>>;
  /** Genuine unsupported effects reject by default in XLSX/PDF; text accepts them with diagnostics. */
  readonly unsupportedPresentation?: "reject" | "text";
  readonly datetime?: "preserve" | "typed" | "text";
  readonly xlsx?: Pick<XlsxOptions, "dateMode">;
  readonly signal?: AbortSignal;
  readonly onDiagnostic?: (diagnostic: CanopyDiagnostic) => void;
}
/** Flat, reusable typed rows. Full writers additionally map captured groups and native rules. */
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
  /** Existing grouped-writer callbacks. Omitted labels use the documented English defaults. */
  readonly grouped?: Partial<Pick<GroupedPdfOptions, "subtotalLabel" | "grandTotalLabel" | "includeGrandTotal" | "continuationLabel" | "groupHeadingLevel">> & {
    readonly labelColumn?: Pick<ColumnSettings, "header">;
  };
  /** Add a metadata worksheet containing only captured export metadata and its accepted query. */
  readonly about?: boolean;
}

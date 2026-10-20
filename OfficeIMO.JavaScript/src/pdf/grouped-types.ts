import type { GroupedReportEvent, GroupedReportSummary } from "../core/grouped-report.js";
import type { ExportValue } from "../core/index.js";
import type { PdfOptions } from "./types.js";

/** Flat analytical groups share the ordinary table's fonts, layout, limits and display formatting. */
export interface GroupedPdfOptions extends Omit<PdfOptions<readonly ExportValue[]>, "columns" | "footer" | "headerRows" | "tagged"> {
  /** Group/table semantic associations are not qualified for tagged output. */
  readonly tagged?: never;
  /** An additional leading text column leaves every original data column and numeric total intact. */
  readonly labelColumn: { readonly header: string; readonly width?: number };
  /** Explicit host-authored label; the event includes the group label and all accepted body-row count. */
  readonly subtotalLabel: (group: Extract<GroupedReportEvent, { readonly kind: "groupEnd" }>) => string;
  /** Explicit host-authored label for the final summary. */
  readonly grandTotalLabel: (summary: GroupedReportSummary) => string;
  /** Include the final summary row, including an empty report's zero/null totals. Default true. */
  readonly includeGrandTotal?: boolean;
  /** Repeated before column headings on continued group pages. Defaults to the exact captured group label. */
  readonly continuationLabel?: (group: Extract<GroupedReportEvent, { readonly kind: "groupStart" }>) => string;
}

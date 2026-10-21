import type { GroupedReportEvent, GroupedReportSummary } from "../core/grouped-report.js";
import type { ExportValue } from "../core/index.js";
import type { PdfOptions } from "./types.js";
import type { PdfTableOptions } from "./document-types.js";

/** Host-authored labels shared by standalone and composed analytical PDF tables. */
interface GroupedPdfLabels {
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

/** A composed analytical table inherits fonts, page geometry, cancellation and cumulative limits from its document. */
export interface GroupedPdfTableOptions extends Omit<PdfTableOptions<readonly ExportValue[]>, "columns" | "footer" | "headerRows">, GroupedPdfLabels {}

/** Analytical groups use the composed document/table layout owner and its resource policy. */
export interface GroupedPdfOptions extends Omit<PdfOptions<readonly ExportValue[]>, "columns" | "footer" | "headerRows" | "tagged">, GroupedPdfLabels {
  /** Group/table semantic associations are not qualified for tagged output. */
  readonly tagged?: never;
}

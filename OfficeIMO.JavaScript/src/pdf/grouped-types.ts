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
  /** Tagged heading level for the outermost group, followed by one level per nested group. Default 2.
   * Every opened group must fit through heading level 6; choose 1 for a titleless standalone report. */
  readonly groupHeadingLevel?: 1 | 2 | 3 | 4 | 5 | 6;
}

/** A composed analytical table inherits fonts, page geometry, cancellation and cumulative limits from its document. */
export interface GroupedPdfTableOptions extends Omit<PdfTableOptions<readonly ExportValue[]>, "columns" | "footer" | "headerRows">, GroupedPdfLabels {}

/** Analytical groups use the composed document/table layout owner and its resource policy. */
export interface GroupedPdfOptions extends Omit<PdfOptions<readonly ExportValue[]>, "columns" | "footer" | "headerRows">, GroupedPdfLabels {}

import { assertGroupedReportSource } from "../core/grouped-report.js";
import type { GroupedReportSource } from "../core/grouped-report.js";
import type { PdfGroupedTableBlock } from "./document-types.js";
import type { GroupedPdfTableOptions } from "./grouped-types.js";
import { captureFields } from "../internal/capture.js";
import { captureWideTable } from "./settings.js";
import { presentation } from "./presentation-settings.js";

/** @internal Capture schemas shared by the block factory and standalone convenience writers. */
export const groupedPdfLabelFields = ["labelColumn", "subtotalLabel", "grandTotalLabel", "includeGrandTotal", "continuationLabel"] as const;
/** @internal A grouped table may override presentation but not its document's resource/lifetime policy. */
export const groupedPdfTableFields = ["includeHeader", "headerPresentation", "footerPresentation", "alternateRowColor", "columnWidths", "wideTable", "formatValue", "formatRuns", ...groupedPdfLabelFields] as const;

/** Capture an analytical table without opening its row iterator or attaching cancellation listeners.
 * Ownership begins when the document reaches the block; unused blocks leave their sources unconsumed. */
export function pdfGroupedTable(report: GroupedReportSource, options: GroupedPdfTableOptions): PdfGroupedTableBlock {
  assertGroupedReportSource(report);
  return Object.freeze({ kind: "groupedTable", report, options: captureGroupedPdfTableOptions(options) });
}

/** @internal Capture the same public table contract for factory-created and structural document blocks. */
export function captureGroupedPdfTableOptions(configuration: GroupedPdfTableOptions): GroupedPdfTableOptions {
  if (!configuration || typeof configuration !== "object" || Array.isArray(configuration) ||
    Object.keys(configuration).some(key => !(groupedPdfTableFields as readonly string[]).includes(key)))
    throw new TypeError("Grouped PDF table options support table presentation and labels; document settings belong on writePdfDocument.");
  const captured = captureFields(configuration, groupedPdfTableFields);
  const { labelColumn, subtotalLabel, grandTotalLabel, includeGrandTotal, continuationLabel } = captured;
  if (!labelColumn || typeof labelColumn !== "object") throw new TypeError("Grouped PDF needs a leading labelColumn.");
  const { header, width } = labelColumn;
  if (typeof header !== "string") throw new TypeError("The grouped label column header must be text.");
  for (const [name, value] of [["subtotalLabel", subtotalLabel], ["grandTotalLabel", grandTotalLabel]] as const)
    if (typeof value !== "function") throw new TypeError(name + " must be a synchronous text callback.");
  if (continuationLabel !== undefined && typeof continuationLabel !== "function") throw new TypeError("continuationLabel must be a synchronous text callback.");
  if (includeGrandTotal !== undefined && typeof includeGrandTotal !== "boolean") throw new TypeError("includeGrandTotal must be boolean.");
  return Object.freeze({ ...captured, ...(captured.wideTable === undefined ? {} : { wideTable: captureWideTable(captured.wideTable)! }), labelColumn: Object.freeze({ header, ...(width === undefined ? {} : { width }) }),
    ...(captured.columnWidths === undefined ? {} : { columnWidths: Object.freeze([...captured.columnWidths]) }),
    ...(captured.headerPresentation === undefined ? {} : { headerPresentation: presentation(captured.headerPresentation) }),
    ...(captured.footerPresentation === undefined ? {} : { footerPresentation: presentation(captured.footerPresentation) }) });
}

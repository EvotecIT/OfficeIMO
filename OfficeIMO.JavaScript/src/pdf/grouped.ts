import { assertGroupedReportSource } from "../core/grouped-report.js";
import type { GroupedReportSource } from "../core/grouped-report.js";
import { BlobByteSink } from "../core/sinks.js";
import type { ExportResult, OutputDestination } from "../core/index.js";
import { writePdfDocumentTo } from "./document.js";
import type { PdfDocumentOptions } from "./document-types.js";
import type { GroupedPdfOptions, GroupedPdfTableOptions } from "./grouped-types.js";
import { pdfOptionFields } from "./settings.js";
import { captureFields } from "../internal/capture.js";
import { pdfGroupedTable, groupedPdfLabelFields, groupedPdfTableFields } from "./grouped-table.js";
export { pdfGroupedTable } from "./grouped-table.js";

const groupedOptionFields = [...pdfOptionFields, ...groupedPdfLabelFields] as readonly (keyof GroupedPdfOptions)[];

/** Stream analytical groups through the composed document owner. No group body is retained. */
export async function writeGroupedPdf(report: GroupedReportSource, configuration: GroupedPdfOptions): Promise<Blob> {
  const sink = new BlobByteSink();
  try { await writeGroupedPdfTo(report, sink, configuration); return sink.toBlob("application/pdf"); }
  catch (error) { sink.discard(); throw error; }
}

/** Result/progress rows count source data. Generated subtotals and grand totals also consume maxRows.
 * Tagged group associations are unsupported; oversized detail/subtotal bundles use ordinary row pagination. */
export async function writeGroupedPdfTo(report: GroupedReportSource, destination: OutputDestination,
  configuration: GroupedPdfOptions): Promise<ExportResult> {
  assertGroupedReportSource(report);
  if (!configuration || typeof configuration !== "object") throw new TypeError("Grouped PDF options are required.");
  const captured = captureFields(configuration, groupedOptionFields);
  if (captured.tagged !== undefined) throw new TypeError("Tagged grouped PDF reports are unsupported.");
  const tableOptions = Object.fromEntries(groupedPdfTableFields.filter(name => Object.hasOwn(captured, name)).map(name => [name, captured[name]])) as unknown as GroupedPdfTableOptions;
  const documentOptions = { ...captured } as Record<string, unknown>;
  for (const name of groupedPdfTableFields) delete documentOptions[name];
  delete documentOptions.columns; delete documentOptions.headerRows; delete documentOptions.footer;
  return writePdfDocumentTo([pdfGroupedTable(report, tableOptions)], destination, documentOptions as PdfDocumentOptions);
}

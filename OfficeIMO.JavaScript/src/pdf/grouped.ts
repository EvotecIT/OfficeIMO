import { assertGroupedReportSource, groupedReportOperation } from "../core/grouped-report.js";
import type { GroupedReportSource, GroupedReportSummary } from "../core/grouped-report.js";
import { BlobByteSink } from "../core/sinks.js";
import { notifyProgress } from "../core/progress.js";
import { ExportCell } from "../core/presentation.js";
import type { Column, ColumnValueContext, ExportValue, ExportResult, OutputDestination } from "../core/index.js";
import { pdfTable, writePdfDocumentTo } from "./document.js";
import type { PdfDocumentBlock, PdfDocumentOptions, PdfTableOptions } from "./document-types.js";
import type { GroupedPdfOptions } from "./grouped-types.js";
import { settings } from "./settings.js";
import { synchronousText } from "./text.js";

/** Stream contiguous analytical groups through the ordinary PDF document/table writer. No group body is retained. */
export async function writeGroupedPdf(report: GroupedReportSource, configuration: GroupedPdfOptions): Promise<Blob> {
  const sink = new BlobByteSink();
  try { await writeGroupedPdfTo(report, sink, configuration); return sink.toBlob("application/pdf"); }
  catch (error) { sink.discard(); throw error; }
}

/** Rows/progress count source data only. Generated labels/subtotals still consume ordinary PDF row/cell/text limits.
 * Tagged grouped reports are unsupported; leaf-table tagging alone does not qualify group semantics. */
export async function writeGroupedPdfTo(report: GroupedReportSource, destination: OutputDestination,
  configuration: GroupedPdfOptions): Promise<ExportResult> {
  assertGroupedReportSource(report);
  if (!configuration || typeof configuration !== "object") throw new TypeError("Grouped PDF options are required.");
  const { labelColumn, subtotalLabel, grandTotalLabel, includeGrandTotal = true, continuationLabel, ...input } = configuration;
  if (!labelColumn || typeof labelColumn !== "object") throw new TypeError("Grouped PDF needs a leading labelColumn.");
  const { header, width } = labelColumn;
  if (typeof header !== "string") throw new TypeError("The grouped label column header must be text.");
  for (const [name, value] of [["subtotalLabel", subtotalLabel], ["grandTotalLabel", grandTotalLabel]] as const)
    if (typeof value !== "function") throw new TypeError(name + " must be a synchronous text callback.");
  if (continuationLabel !== undefined && typeof continuationLabel !== "function") throw new TypeError("continuationLabel must be a synchronous text callback.");
  if (typeof includeGrandTotal !== "boolean") throw new TypeError("includeGrandTotal must be boolean.");
  if (input.tagged !== undefined) throw new TypeError("Tagged grouped PDF reports are unsupported.");
  const operation = groupedReportOperation(report, input.signal);
  try {
  const columns = Object.freeze([Object.freeze({ header, ...(width === undefined ? {} : { width }) }), ...report.columns]);
  // Reuse the normal capture/validation owner before consuming any event or invoking display callbacks.
  const captured = settings({ ...input, ...(operation.signal === undefined ? {} : { signal: operation.signal }), columns }).options;
  const { columns: _columns, includeHeader, headerRows: _headers, footer: _footer, headerPresentation, footerPresentation,
    alternateRowColor, columnWidths, wideTable, formatValue, formatRuns, onProgress, ...documentOptions } = captured;
  let dataRows = 0, currentRowIndex = -1, summary: GroupedReportSummary | undefined;
  const sourceContext = (context: ColumnValueContext): ColumnValueContext => ({ rowIndex: currentRowIndex,
    columnIndex: context.columnIndex - 1, column: report.columns[context.columnIndex - 1]! });
  const tableOptions: PdfTableOptions<readonly ExportValue[]> = {
    columns, ...(includeHeader === undefined ? {} : { includeHeader }),
    ...(headerPresentation === undefined ? {} : { headerPresentation }),
    ...(alternateRowColor === undefined ? {} : { alternateRowColor }),
    ...(columnWidths === undefined ? {} : { columnWidths }), ...(wideTable === undefined ? {} : { wideTable }),
    ...(formatValue === undefined ? {} : { formatValue: (value, context) => formatValue(value, sourceContext(context)) }),
    ...(formatRuns === undefined ? {} : { formatRuns: (value, context) => formatRuns(value, sourceContext(context)) })
  };
  const summaryValues = (label: string, values: readonly (number | null)[]): readonly ExportValue[] => Object.freeze([
    new ExportCell(label, { text: label, presentation: footerPresentation ?? { bold: true, background: "eef2f6" } }),
    ...values.map(value => new ExportCell(value, { presentation: footerPresentation ?? { bold: true, background: "eef2f6" } }))
  ]);
  async function* blocks(): AsyncGenerator<PdfDocumentBlock> {
    const cursor = operation.events[Symbol.asyncIterator]();
    try {
      while (true) {
        const item = await cursor.next();
        if (item.done) break;
        const event = item.value;
        if (event.kind === "groupStart") {
          const continued = continuationLabel === undefined ? event.label : synchronousText(continuationLabel(event), "continuationLabel");
          async function* rows(): AsyncGenerator<readonly ExportValue[]> {
            while (true) {
              const next = await cursor.next();
              if (next.done) throw new TypeError("Grouped report ended before its group summary.");
              const value = next.value;
              if (value.kind === "row") {
                currentRowIndex = value.rowIndex; dataRows = value.rowIndex + 1;
                yield [new ExportCell("", { text: "" }), ...value.values];
              } else if (value.kind === "groupEnd") {
                currentRowIndex = -1;
                yield summaryValues(synchronousText(subtotalLabel(value), "subtotalLabel"), value.totals);
                return;
              } else throw new TypeError("Unexpected grouped report event inside a group.");
            }
          }
          yield { ...pdfTable(rows(), tableOptions),
            caption: { text: event.label, presentation: { bold: true }, spaceAfter: 4 },
            continuationCaption: { text: continued, presentation: { bold: true }, spaceAfter: 4 } };
        } else if (event.kind === "complete") {
          summary = event;
          if (includeGrandTotal || event.rows === 0) {
            currentRowIndex = -1;
            // A heading-free grand summary has no group-heading hierarchy. Keep the original
            // source metadata for formatting contexts and for an empty report's visible schema.
            const completeColumns = event.rows === 0 ? columns : columns.map(column => {
              const { groups: _groups, ...leaf } = column as Column<readonly ExportValue[]>;
              return leaf;
            });
            yield pdfTable(includeGrandTotal ? [summaryValues(synchronousText(grandTotalLabel(event), "grandTotalLabel"), event.totals)] : [],
              { ...tableOptions, columns: completeColumns, includeHeader: event.rows === 0 ? includeHeader ?? true : false });
          }
          break;
        } else throw new TypeError("Unexpected grouped report event outside a group.");
      }
    } finally { await cursor.return?.(); }
  }
  const result = await writePdfDocumentTo(blocks(), destination, { ...documentOptions,
    onProgress: progress => notifyProgress(onProgress, { ...progress, rows: progress.phase === "complete" ? summary?.rows ?? dataRows : dataRows }) } as PdfDocumentOptions);
  return { ...result, rows: summary?.rows ?? dataRows, columns: columns.length };
  } finally { operation.dispose(); }
}

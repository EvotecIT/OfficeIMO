import { BlobByteSink, ExportCell } from "../../core/index.js";
import type { OutputDestination, ExportResult, ExportProgress } from "../../core/index.js";
import { notifyProgress } from "../../core/progress.js";
import { captureFields } from "../../internal/capture.js";
import { captureCommonLimits, captureXlsxOptions, captureCsvOptions, capturePdfOptions, mergeAdapterLimits } from "../options.js";
import { writeXlsxTo } from "../../xlsx/index.js";
import { writeCsvTo } from "../../csv/index.js";
import { writePdfTo, writeGroupedPdfTo } from "../../pdf/index.js";
import { prepareCanopyExport } from "./source.js";
import { canopyPins } from "./presentation.js";
import { canopyRules } from "./rules.js";
import { canopyGrouped } from "./grouping.js";
import { canopyAbout } from "./metadata.js";
import { writeCanopyWorkbook } from "./workbook.js";
import type { CanopyCapture, CanopyExport, CanopyFormat, CanopyWriteOptions } from "./types.js";
export { createCanopyExport } from "./source.js";
export { ExportCell } from "../../core/presentation.js";
export { PdfFont } from "../../pdf/font.js";
export type { CanopyCapture, CanopyExportCell, CanopyExportColumn, CanopyExportRow, CanopyExport, CanopyFormat, CanopyDiagnostic, CanopySourceOptions, CanopyWriteOptions,
  CanopyNumberFormat, CanopyExportGroup, CanopyExportGrouping, CanopyRuleFilter, CanopyRulePredicate, CanopyFormatRule } from "./types.js";

/** Write a captured CanopyX grid to a caller-owned destination; partial bytes remain caller-owned on failure. */
export async function writeCanopyTo(capture: CanopyCapture, format: CanopyFormat, destination: OutputDestination, options: CanopyWriteOptions = {}): Promise<ExportResult> {
  const selected = captureFields(options, ["tones", "columnOptions", "unsupportedPresentation", "datetime", "xlsx", "csv", "pdf", "signal", "onDiagnostic", "limits", "onProgress", "grouped", "about"]);
  const { signal, onProgress } = selected, limits = captureCommonLimits(selected.limits);
  const xlsx = captureXlsxOptions(format === "xlsx" ? selected.xlsx : undefined);
  const csv = captureCsvOptions(format === "csv" ? selected.csv : undefined);
  const pdf = capturePdfOptions(format === "pdf" ? selected.pdf : undefined);
  const prepared = prepareCanopyExport(capture, format, { ...selected, xlsx }), { request, report, tones } = prepared;
  if (selected.about !== undefined && typeof selected.about !== "boolean") throw new TypeError("Canopy about must be boolean.");
  const labels = selected.grouped === undefined ? undefined : captureFields(selected.grouped, ["labelColumn", "subtotalLabel", "grandTotalLabel", "includeGrandTotal", "continuationLabel", "groupHeadingLevel"]);
  const capturedLabels = labels && { ...labels, ...(labels.labelColumn === undefined ? {} : { labelColumn: Object.freeze({ header: labels.labelColumn.header }) }) };
  const about = format === "xlsx" && selected.about ? canopyAbout(request, xlsx) : undefined;
  const ruleCellTypes = new Map<number, "number" | "date">();
  let consumed = 0;
  const source: CanopyExport = Object.freeze({ ...prepared.source, async *rows() {
    for await (const row of prepared.source.rows()) {
      for (const [column, type] of ruleCellTypes) {
        const cell = row[column], raw = cell instanceof ExportCell ? cell.value : cell;
        if (raw == null || type === "number" && typeof raw === "number" && Number.isFinite(raw) || type === "date" && raw instanceof Date && Number.isFinite(raw.getTime())) continue;
        const code = "UNSUPPORTED_NATIVE_RULE_CELL_TYPE:" + type;
        report({ code, columnId: source.columns[column]!.key });
        // Text fallback cannot retract an already captured native declaration after streaming starts.
        throw new TypeError("Unsupported Canopy presentation: " + code + " (" + source.columns[column]!.key + ")");
      }
      consumed++; yield row;
    }
  } });
  const progress = onProgress ? { onProgress: (event: ExportProgress) => {
    // About and generated group rows consume output budgets, but never become exported source details.
    notifyProgress(onProgress, { ...event, rows: consumed,
      ...(event.sheetRows === undefined ? {} : { sheetRows: about !== undefined && event.sheetName === "About" ? 0 : consumed }),
      ...(source.rowCount === null ? {} : { totalRows: source.rowCount }) });
  } } : {};
  const stream = { ...(signal ? { signal } : {}), ...progress };
  if (format === "xlsx") {
    const formatLimits = mergeAdapterLimits(xlsx.limits, limits), grouped = canopyGrouped(source, request, { ...selected, xlsx }, formatLimits, report, format);
    const pinned = canopyPins(request.columns, report), rules = canopyRules(request, format, grouped !== undefined, tones, report, { ...selected, xlsx }, ruleCellTypes);
    const sheet = { ...xlsx.sheet,
      freezeHeader: xlsx.sheet?.freezeHeader ?? (xlsx.sheet?.includeHeader !== false),
      freezeColumns: xlsx.sheet?.freezeColumns ?? (pinned ? pinned + (grouped ? 1 : 0) : 0),
      conditionalFormats: [...rules, ...(xlsx.sheet?.conditionalFormats ?? [])] };
    const settings = { ...xlsx, ...stream, columns: source.columns, sheet, dateMode: xlsx.dateMode ?? request.timeZone, limits: formatLimits };
    if (grouped || about) return writeCanopyWorkbook(source, grouped, destination, settings, capturedLabels, about);
    return writeXlsxTo(source.rows(), destination, settings);
  }
  if (format === "pdf") {
    const formatLimits = mergeAdapterLimits(pdf.limits, limits), grouped = canopyGrouped(source, request, selected, formatLimits, report, format);
    canopyRules(request, format, grouped !== undefined, tones, report, selected);
    // CSS pixels become points before the existing fit/band policy. Explicit PDF point widths win.
    const fitted = request.columns.every(column => column.width !== undefined) ? request.columns.map(column => column.width!.preferred * 72 / 96) : undefined;
    const widths = pdf.columnWidths ?? fitted;
    const settings = { ...(fitted === undefined ? {} : { orientation: "landscape" as const, wideTable: "fit" as const }), ...pdf, ...stream, limits: formatLimits };
    if (grouped) {
      if (pdf.headerRows !== undefined || pdf.footer !== undefined) throw new TypeError("Grouped Canopy PDF owns heading and summary rows.");
      const { headerRows: _headers, footer: _footer, ...groupedSettings } = settings;
      return writeGroupedPdfTo(grouped, destination, { ...groupedSettings,
        ...(widths === undefined ? {} : { columnWidths: [72, ...widths] }),
        labelColumn: { header: capturedLabels?.labelColumn?.header ?? "Group", width: 12 },
        subtotalLabel: capturedLabels?.subtotalLabel ?? (group => group.label + " total"),
        grandTotalLabel: capturedLabels?.grandTotalLabel ?? (() => "Grand total"),
        ...(capturedLabels?.includeGrandTotal === undefined ? {} : { includeGrandTotal: capturedLabels.includeGrandTotal }),
        ...(capturedLabels?.continuationLabel === undefined ? {} : { continuationLabel: capturedLabels.continuationLabel }),
        ...(capturedLabels?.groupHeadingLevel === undefined ? {} : { groupHeadingLevel: capturedLabels.groupHeadingLevel }) });
    }
    return writePdfTo(source.rows(), destination, { ...settings, columns: source.columns, ...(widths === undefined ? {} : { columnWidths: widths }) });
  }
  return writeCsvTo(source.rows(), destination, { ...csv, ...stream, columns: source.columns, valueMode: request.values,
    limits: mergeAdapterLimits(csv.limits, limits) });
}

/** Export an immutable native capture as a Blob without loading or registering CanopyX. */
export async function exportCanopy(capture: CanopyCapture, format: CanopyFormat, options: CanopyWriteOptions = {}): Promise<Blob> {
  const sink = new BlobByteSink();
  try {
    await writeCanopyTo(capture, format, sink, options);
    return sink.toBlob(format === "xlsx" ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" : format === "pdf" ? "application/pdf" : "text/csv;charset=utf-8");
  } catch (error) { sink.discard(); throw error; }
}

import { BlobByteSink, checkAbort, saveBlob } from "../../core/index.js";
import type { OutputDestination, ExportResult } from "../../core/index.js";
import { notifyProgress } from "../../core/progress.js";
import { writeCsvTo } from "../../csv/index.js";
import { writeXlsxTo } from "../../xlsx/index.js";
import { writePdfTo } from "../../pdf/index.js";
import { portableSheet, portableWorkbook } from "../../xlsx/portable.js";
import { call, member } from "./api.js";
import { value } from "./headings.js";
import { captureDataTablesExport } from "./snapshot.js";
import { captureFields } from "../../internal/capture.js";
import { captureCommonLimits, captureXlsxOptions, captureCsvOptions, capturePdfOptions, mergeAdapterLimits } from "../options.js";
import type { DataTablesApi, DataTablesButtonOptions, DataTablesHost, DataTablesWriteOptions } from "./types.js";
export { createDataTablesExport } from "./snapshot.js";
export { ExportCell } from "../../core/presentation.js";
export { PdfFont } from "../../pdf/font.js";
export type { DataTablesApi, DataTablesHost, DataTablesOptions, DataTablesExportOptions, DataTablesFormat,
  DataTablesCellContext, DataTablesExport, DataTablesWriteOptions, DataTablesButtonOptions } from "./types.js";
const writeOptionFields = ["mode", "exportOptions", "batchRows", "maxBatchCells", "headings", "includeFooter", "serverSide", "columnOptions", "project", "signal", "limits", "onProgress", "xlsx", "csv", "pdf"] as const;
const buttonOptionFields = [...writeOptionFields, "filename", "save", "onError"] as const;

/** Write directly to a caller-owned destination. Failed destinations own disposal of their partial bytes. */
export async function writeDataTableTo(host: DataTablesHost, table: DataTablesApi, format: "xlsx" | "csv" | "pdf", destination: OutputDestination,
  options: DataTablesWriteOptions = {}): Promise<ExportResult> {
  if (format !== "xlsx" && format !== "csv" && format !== "pdf") throw new TypeError("DataTables export format must be xlsx, csv or pdf.");
  for (const key of ["workbook", "sheet", "sheetName"])
    if (key in options) throw new TypeError("DataTables writer settings belong in xlsx; use xlsx.sheet.name for the worksheet name.");
  const captured = captureFields(options, writeOptionFields);
  const { signal, onProgress } = captured, limits = captureCommonLimits(captured.limits);
  const xlsx = captureXlsxOptions(format === "xlsx" ? captured.xlsx : undefined);
  const csv = captureCsvOptions(format === "csv" ? captured.csv : undefined);
  const pdf = capturePdfOptions(format === "pdf" ? captured.pdf : undefined);
  const sheet = xlsx.sheet;
  if (format === "xlsx") { portableSheet(sheet); portableWorkbook(xlsx); }
  const replacementXlsxFooter = sheet?.footer && [sheet.footer.rows, sheet.footer.values, sheet.footer.totals].some(value => value !== undefined);
  const source = captureDataTablesExport(host, table, { ...captured, ...(limits === undefined ? {} : { limits }), ...(format === "pdf" ? { headings: captured.headings ?? "structured" } : {}) }, {
    header: format === "pdf" ? pdf.includeHeader === false || pdf.headerRows !== undefined
      : format === "xlsx" ? sheet?.includeHeader === false || sheet?.headerRows !== undefined
        : csv.includeHeader === false || csv.headerRows !== undefined,
    footer: format === "pdf" ? pdf.footer !== undefined : format === "xlsx" ? Boolean(replacementXlsxFooter) : csv.footer !== undefined
  });
  let result: ExportResult;
  const stream = { ...(signal ? { signal } : {}), ...(onProgress ? {
    onProgress: (event: import("../../core/index.js").ExportProgress) => notifyProgress(onProgress, { ...event, totalRows: source.rowCount })
  } : {}) };
  if (format === "xlsx") {
    const sourceFooter = source.footerStructure ? { rows: source.footerStructure } : source.footer ? { values: source.footer } : undefined;
    const footer = replacementXlsxFooter ? sheet?.footer : sourceFooter || sheet?.footer ? { ...sourceFooter, ...sheet?.footer } : undefined;
    result = await writeXlsxTo(source.rows, destination, { ...xlsx, ...stream, columns: source.columns,
      limits: mergeAdapterLimits(xlsx.limits, limits),
      sheet: { ...sheet, ...(source.headerStructure && sheet?.includeHeader !== false && sheet?.headerRows === undefined ? { headerRows: source.headerStructure } : {}), ...(footer ? { footer } : {}) } });
  } else if (format === "pdf") {
    const footer = pdf.footer ?? (source.footerStructure ? { rows: source.footerStructure } : source.footer ? { values: source.footer } : undefined);
    result = await writePdfTo(source.rows, destination, { ...pdf, ...stream, columns: source.columns,
      limits: mergeAdapterLimits(pdf.limits, limits),
      ...(source.headerStructure && pdf.includeHeader !== false && !pdf.headerRows ? { headerRows: source.headerStructure } : {}),
      ...(footer ? { footer } : {}) });
  } else {
    const footer = csv.footer ?? (source.footerStructure ? { rows: source.footerStructure }
      : source.footer ? { rows: [source.footer.map(value => ({ value }))] } : undefined);
    result = await writeCsvTo(source.rows, destination, { ...csv, ...stream, columns: source.columns,
      limits: mergeAdapterLimits(csv.limits, limits),
      ...(csv.includeHeader !== false && csv.headerRows === undefined ? { headerRows: source.headerStructure ?? source.headers.map(row => row.map(value => ({ value }))) } : {}), ...(footer ? { footer } : {}) });
  }
  checkAbort(signal);
  return { rows: source.rowCount, columns: source.columns.length, bytes: result.bytes };
}

/** Return a browser/worker/Node Blob. Use writeDataTableTo to avoid retaining the completed file in memory. */
export async function exportDataTable(host: DataTablesHost, table: DataTablesApi, format: "xlsx" | "csv" | "pdf", options: DataTablesWriteOptions = {}): Promise<Blob> {
  const sink = new BlobByteSink();
  try {
    await writeDataTableTo(host, table, format, sink, options);
    return sink.toBlob(format === "pdf" ? "application/pdf" : format === "csv" ? "text/csv;charset=utf-8" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  } catch (error) { sink.discard(); throw error; }
}

/** Register opt-in Buttons entries; no import-time registration, DataTables import or jQuery dependency. */
export function registerDataTablesButtons(host: DataTablesHost, options: DataTablesButtonOptions = {}): void {
  const buttons = host.ext?.buttons;
  if (!buttons || typeof member(host.Buttons, "stripData") !== "function") throw new TypeError("Install DataTables Buttons before registering OfficeIMO exports.");
  if (buttons.officeimoExcel || buttons.officeimoCsv || buttons.officeimoPdf) throw new TypeError("OfficeIMO DataTables buttons are already registered.");
  for (const [name, format, label] of [["officeimoExcel", "xlsx", "Excel"], ["officeimoCsv", "csv", "CSV"], ["officeimoPdf", "pdf", "PDF"]] as const) {
    buttons[name] = { text: label, async: 1,
      action: function (_event: unknown, table: DataTablesApi, _node: unknown, configuration: unknown, complete?: () => void): void {
        let current = options;
        const finish = () => { const callback = complete; complete = undefined; callback?.(); };
        void (async () => {
          try {
            const configured = member(configuration, "officeimo"), filename = member(configuration, "filename");
            if (configured !== undefined && (!configured || typeof configured !== "object")) throw new TypeError("Button officeimo settings must be an object.");
            const defaults = captureFields(options, buttonOptionFields), overrides = captureFields((configured ?? {}) as DataTablesButtonOptions, buttonOptionFields);
            current = { ...defaults, ...overrides,
              ...(filename !== undefined ? { filename: filename as NonNullable<DataTablesButtonOptions["filename"]> } : {}),
              ...(member(configuration, "footer") === false ? { includeFooter: false } : {}),
              ...(member(configuration, "exportOptions") ? { exportOptions: member(configuration, "exportOptions") as NonNullable<DataTablesButtonOptions["exportOptions"]> } : {}) };
            if (member(configuration, "customize") !== undefined) throw new TypeError("Native customize callbacks are unsupported; use OfficeIMO xlsx/pdf options.");
            if (format === "xlsx") {
              const base = captureXlsxOptions(defaults.xlsx), selected = captureXlsxOptions(overrides.xlsx);
              current = { ...current, xlsx: { ...base, ...selected,
                ...(base.sheet || selected.sheet ? { sheet: { ...base.sheet, ...selected.sheet } } : {}) } };
            } else if (format === "csv") current = { ...current, csv: { ...captureCsvOptions(defaults.csv), ...captureCsvOptions(overrides.csv) } };
            if (format === "pdf") {
              const info = call(table.buttons, "exportInfo", configuration);
              const pdf = { ...capturePdfOptions(defaults.pdf), ...capturePdfOptions(overrides.pdf) };
              for (const key of ["title", "messageTop", "messageBottom"] as const) if (member(configuration, key) !== undefined) {
                if (member(configuration, key) === null) { delete pdf[key]; continue; }
                const text = member(info, key);
                if (text === null || text === "") { delete pdf[key]; continue; }
                if (typeof text !== "string") throw new TypeError("PDF " + key + " must resolve to text.");
                pdf[key] = text;
              }
              for (const key of ["orientation", "pageSize"] as const) if (member(configuration, key) !== undefined)
                Object.assign(pdf, { [key]: member(configuration, key) });
              if (member(configuration, "header") === false) { pdf.includeHeader = false; delete pdf.headerRows; }
              if (member(configuration, "footer") === false) delete pdf.footer;
              current = { ...current, pdf };
            } else if (member(configuration, "header") === false || ["title", "messageTop", "messageBottom"].some(key => member(configuration, key) != null))
              throw new TypeError("Use OfficeIMO xlsx.sheet title/footer options; native Buttons report layout options are unsupported.");
            const pattern = value(typeof current.filename === "function" ? current.filename(configuration, table) : current.filename ?? "Export");
            if (typeof pattern !== "string" || !pattern.trim()) throw new TypeError("Export filename must be a non-empty string.");
            const file = value(member(call(table.buttons, "exportInfo", { filename: pattern, extension: "" }), "filename"));
            if (typeof file !== "string" || !file.trim()) throw new TypeError("Resolved export filename must be a non-empty string.");
            const blob = await exportDataTable(host, table, format, current);
            await (current.save ?? saveBlob)(blob, file.toLowerCase().endsWith("." + format) ? file : file + "." + format);
          } catch (error) {
            // Completion runs before the error handler so application error UI cannot strand Buttons' spinner.
            finish();
            if (current.onError) await current.onError(error);
            else console.error("OfficeIMO table export failed.", error);
          } finally { finish(); }
        })().catch(error => console.error("OfficeIMO export error handler failed.", error));
      }
    };
  }
}

import { BlobByteSink, checkAbort, saveBlob } from "../../core/index.js";
import type { OutputDestination, ExportResult, ExportValue } from "../../core/index.js";
import { writeCsvTo } from "../../csv/index.js";
import { writeXlsxTo } from "../../xlsx/index.js";
import { writePdfTo } from "../../pdf/index.js";
import { portableSheet, portableWorkbook } from "../../xlsx/portable.js";
import { call, member } from "./api.js";
import { value } from "./headings.js";
import { createDataTablesExport } from "./snapshot.js";
import type { DataTablesApi, DataTablesButtonOptions, DataTablesHost, DataTablesWriteOptions } from "./types.js";
export { createDataTablesExport } from "./snapshot.js";
export { ExportCell } from "../../core/presentation.js";
export type { DataTablesApi, DataTablesHost, DataTablesOptions, DataTablesExportOptions, DataTablesFormat,
  DataTablesCellContext, DataTablesExport, DataTablesWriteOptions, DataTablesButtonOptions } from "./types.js";

/** Write directly to a caller-owned destination. Failed destinations own disposal of their partial bytes. */
export async function writeDataTableTo(host: DataTablesHost, table: DataTablesApi, format: "xlsx" | "csv" | "pdf", destination: OutputDestination,
  options: DataTablesWriteOptions = {}): Promise<ExportResult> {
  if (format !== "xlsx" && format !== "csv" && format !== "pdf") throw new TypeError("DataTables export format must be xlsx, csv or pdf.");
  if (format !== "pdf" && options.headings === "structured") throw new TypeError("Structured headings require PDF output; use grouped or leaf for Excel/CSV.");
  portableSheet(options.sheet); portableWorkbook(options.workbook);
  const source = createDataTablesExport(host, table, format === "pdf" ? { ...options, headings: options.headings ?? "structured" } : options), signal = options.signal;
  let result: ExportResult;
  const stream = { ...(signal ? { signal } : {}), ...(options.limits ? { limits: options.limits } : {}) };
  if (format === "xlsx") {
    const footer = source.footer || options.sheet?.footer ? { values: source.footer ?? [], ...options.sheet?.footer } : undefined;
    result = await writeXlsxTo(source.rows, destination, { ...options.workbook, ...stream, columns: source.columns,
      sheet: { ...options.sheet, name: options.sheetName ?? "Data", ...(footer ? { footer } : {}) },
      ...(options.onProgress ? { onProgress: event => options.onProgress?.({ ...event, totalRows: source.rowCount }) } : {}) });
  } else if (format === "pdf") {
    const footer = options.pdf?.footer ?? (source.footerStructure ? { rows: source.footerStructure } : source.footer ? { values: source.footer } : undefined);
    result = await writePdfTo(source.rows, destination, { ...options.pdf, ...stream, columns: source.columns,
      ...(source.headerStructure && options.pdf?.includeHeader !== false && !options.pdf?.headerRows ? { headerRows: source.headerStructure } : {}),
      ...(footer ? { footer } : {}),
      ...(options.onProgress ? { onProgress: event => options.onProgress?.({ ...event, totalRows: source.rowCount }) } : {}) });
  } else {
    const headingCount = source.headers.length, footerCount = source.footer ? 1 : 0;
    async function* rows(): AsyncGenerator<readonly ExportValue[]> {
      for (const header of source.headers) yield header;
      yield* source.rows;
      if (source.footer) yield source.footer;
    }
    result = await writeCsvTo(rows(), destination, { ...options.csv, ...stream, columns: source.columns, includeHeader: false,
      ...(options.limits?.maxRows !== undefined ? { limits: { ...options.limits, maxRows: Math.min(Number.MAX_SAFE_INTEGER, options.limits.maxRows + headingCount + footerCount) } } : {}),
      onProgress: event => {
        const rowCount = Math.max(0, Math.min(source.rowCount, event.rows - headingCount));
        options.onProgress?.({ ...event, rows: rowCount, totalRows: source.rowCount });
      } });
  }
  checkAbort(signal);
  return { rows: source.rowCount, columns: source.columns.length, bytes: result.bytes };
}

/** Return a browser/worker/Node Blob. Use writeDataTableTo to avoid retaining the completed file in memory. */
export async function exportDataTable(host: DataTablesHost, table: DataTablesApi, format: "xlsx" | "csv" | "pdf", options: DataTablesWriteOptions = {}): Promise<Blob> {
  const sink = new BlobByteSink();
  try {
    await writeDataTableTo(host, table, format, sink, options);
    checkAbort(options.signal);
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
            current = { ...options, ...(configured as DataTablesButtonOptions | undefined),
              ...(filename !== undefined ? { filename: filename as NonNullable<DataTablesButtonOptions["filename"]> } : {}),
              ...(member(configuration, "footer") === false ? { includeFooter: false } : {}),
              ...(member(configuration, "exportOptions") ? { exportOptions: member(configuration, "exportOptions") as NonNullable<DataTablesButtonOptions["exportOptions"]> } : {}) };
            if (member(configuration, "customize") !== undefined) throw new TypeError("Native customize callbacks are unsupported; use OfficeIMO sheet/workbook/pdf options.");
            if (format === "pdf") {
              const info = call(table.buttons, "exportInfo", configuration);
              const pdf = { ...current.pdf };
              for (const key of ["title", "messageTop", "messageBottom"] as const) if (member(configuration, key) !== undefined) {
                if (member(configuration, key) === null) { delete pdf[key]; continue; }
                const text = member(info, key);
                if (typeof text !== "string") throw new TypeError("PDF " + key + " must resolve to text.");
                pdf[key] = text;
              }
              for (const key of ["orientation", "pageSize"] as const) if (member(configuration, key) !== undefined)
                Object.assign(pdf, { [key]: member(configuration, key) });
              if (member(configuration, "header") === false) pdf.includeHeader = false;
              current = { ...current, pdf };
            } else if (member(configuration, "header") === false || ["title", "messageTop", "messageBottom"].some(key => member(configuration, key) != null))
              throw new TypeError("Use OfficeIMO sheet title/footer options; native Buttons report layout options are unsupported.");
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

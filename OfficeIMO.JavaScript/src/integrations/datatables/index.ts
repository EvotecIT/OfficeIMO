import { BlobByteSink, checkAbort, saveBlob } from "../../core/index.js";
import type { ByteSink, ExportValue } from "../../core/index.js";
import { writeCsvTo } from "../../csv/index.js";
import { Workbook } from "../../xlsx/index.js";
import { member } from "./api.js";
import { value } from "./headings.js";
import { createDataTablesExport } from "./snapshot.js";
import type { DataTablesApi, DataTablesButtonOptions, DataTablesHost, DataTablesWriteOptions } from "./types.js";
export { createDataTablesExport } from "./snapshot.js";
export { ExportCell } from "../../core/presentation.js";
export type { DataTablesApi, DataTablesHost, DataTablesOptions, DataTablesExportOptions, DataTablesFormat,
  DataTablesCellContext, DataTablesExport, DataTablesWriteOptions, DataTablesButtonOptions } from "./types.js";

/** Write directly to a caller-owned destination. Failed destinations own disposal of their partial bytes. */
export async function writeDataTableTo(host: DataTablesHost, table: DataTablesApi, format: "xlsx" | "csv", sink: ByteSink,
  options: DataTablesWriteOptions = {}): Promise<{ readonly rows: number; readonly columns: number; readonly bytes: number }> {
  if (format !== "xlsx" && format !== "csv") throw new TypeError("DataTables export format must be xlsx or csv.");
  const source = createDataTablesExport(host, table, options), signal = options.signal;
  let bytes = 0;
  const destination: ByteSink = { async write(chunk) { await sink.write(chunk); bytes += chunk.byteLength; } };
  const stream = { ...(signal ? { signal } : {}), ...(options.limits ? { limits: options.limits } : {}) };
  if (format === "xlsx") {
    const book = new Workbook({ ...options.workbook, ...stream, sink: destination, ...(options.onProgress ? { onProgress: options.onProgress } : {}) });
    try {
      const footer = source.footer || options.sheet?.footer ? { values: source.footer ?? [], ...options.sheet?.footer } : undefined;
      const sheet = book.addSheet(options.sheetName ?? "Data", { boldHeader: true, autoFilter: true,
        autoSize: { sampleRows: 100, minWidth: 6, maxWidth: 54 }, ...options.sheet, columns: source.columns,
        ...(footer ? { footer } : {}) });
      await sheet.addRows(source.rows); await book.finish();
    } catch (error) { await book.discard(error); throw error; }
  } else {
    const headingCount = source.headers.length, footerCount = source.footer ? 1 : 0;
    async function* rows(): AsyncGenerator<readonly ExportValue[]> {
      for (const header of source.headers) yield header;
      yield* source.rows;
      if (source.footer) yield source.footer;
    }
    await writeCsvTo(rows(), destination, { ...options.csv, ...stream, columns: source.columns, includeHeader: false,
      ...(options.limits?.maxRows !== undefined ? { limits: { ...options.limits, maxRows: Math.min(Number.MAX_SAFE_INTEGER, options.limits.maxRows + headingCount + footerCount) } } : {}),
      onProgress: event => {
        const rowCount = Math.max(0, Math.min(source.rowCount, event.rows - headingCount));
        options.onProgress?.({ ...event, rows: rowCount, ...(event.phase === "complete" ? { bytes } : {}) });
      } });
  }
  checkAbort(signal);
  return { rows: source.rowCount, columns: source.columns.length, bytes };
}

/** Return a browser/worker/Node Blob. Use writeDataTableTo to avoid retaining the completed file in memory. */
export async function exportDataTable(host: DataTablesHost, table: DataTablesApi, format: "xlsx" | "csv", options: DataTablesWriteOptions = {}): Promise<Blob> {
  const sink = new BlobByteSink();
  try {
    await writeDataTableTo(host, table, format, sink, options);
    checkAbort(options.signal);
    return sink.toBlob(format === "csv" ? "text/csv;charset=utf-8" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  } catch (error) { sink.discard(); throw error; }
}

/** Register opt-in Buttons entries; no import-time registration, DataTables import or jQuery dependency. */
export function registerDataTablesButtons(host: DataTablesHost, options: DataTablesButtonOptions = {}): void {
  const buttons = host.ext?.buttons;
  if (!buttons || typeof member(host.Buttons, "stripData") !== "function") throw new TypeError("Install DataTables Buttons before registering OfficeIMO exports.");
  if (buttons.officeimoExcel || buttons.officeimoCsv) throw new TypeError("OfficeIMO DataTables buttons are already registered.");
  for (const [name, format, label] of [["officeimoExcel", "xlsx", "Excel"], ["officeimoCsv", "csv", "CSV"]] as const) {
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
            if (member(configuration, "customize") !== undefined) throw new TypeError("XML customize callbacks are unsupported; use OfficeIMO sheet/workbook options.");
            if (member(configuration, "header") === false || ["title", "messageTop", "messageBottom"].some(key => member(configuration, key) != null))
              throw new TypeError("Use OfficeIMO sheet title/footer options; native Buttons report layout options are unsupported.");
            const blob = await exportDataTable(host, table, format, current);
            const file = value(typeof current.filename === "function" ? current.filename() : current.filename ?? "Export");
            if (typeof file !== "string" || !file.trim()) throw new TypeError("Export filename must be a non-empty string.");
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

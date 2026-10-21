import type { ColumnSettings, ExportProgress, ExportResult, OutputDestination } from "../core/index.js";
import type { GroupedReportEvent, GroupedReportSource } from "../core/grouped-report.js";
import { assertGroupedReportSource, groupedReportOperation } from "../core/grouped-report.js";
import { withDestination } from "../core/sinks.js";
import type { ExportValue } from "../core/presentation.js";
import { Workbook } from "./workbook.js";
import { FormulaCell } from "./formula.js";
import { columnName } from "./values.js";
import { copyStylePatch } from "./styles.js";
import { portableSheet } from "./portable.js";
import type { PortableCellStyle, PortableSheetOptions } from "./portable.js";
import type { WorkbookOptions, XlsxColumn } from "./types.js";
import { captureFields } from "../internal/capture.js";
import { sheetOptionFields, workbookOptionFields } from "./options.js";

type GroupEnd = Extract<GroupedReportEvent, { readonly kind: "groupEnd" }>;
type Complete = Extract<GroupedReportEvent, { readonly kind: "complete" }>;
type GroupedSheet = Omit<PortableSheetOptions, "columns" | "headerRows" | "footer" | "table" | "autoFilter" | "mergedCells" | "hyperlinks" | "conditionalFormats" | "dataValidation" | "rowStyle" | "cellStyle" | "rowOutline" | "maxOutlineLevel">;

/** A grouped workbook owns its leading label column and native summary rows. */
export interface GroupedXlsxOptions extends Omit<WorkbookOptions, "sink" | "cellValueWriters"> {
  /** Separate labels preserve every typed source column, including a numeric first column. */
  readonly labelColumn: Pick<ColumnSettings, "header" | "width" | "wrapText" | "alignment">;
  readonly subtotalLabel: (group: GroupEnd) => string;
  readonly grandTotalLabel: (summary: Complete) => string;
  readonly includeGrandTotal?: boolean;
  /** Hidden details remain included in SUBTOTAL calculations; native filtering is not enabled. */
  readonly collapseGroups?: boolean;
  readonly sheet?: GroupedSheet & { readonly name?: string };
  readonly groupStyle?: PortableCellStyle;
  readonly subtotalStyle?: PortableCellStyle;
  readonly grandTotalStyle?: PortableCellStyle;
}

const unsupported = ["columns", "headerRows", "footer", "table", "mergedCells", "hyperlinks", "conditionalFormats", "dataValidation", "rowStyle", "cellStyle", "rowOutline", "maxOutlineLevel"] as const;
const subtotalOperations = { average: 1, count: 2, max: 4, min: 5, sum: 9 } as const;

function prepare(report: GroupedReportSource, configuration: GroupedXlsxOptions): GroupedXlsxOptions {
  assertGroupedReportSource(report);
  if (!configuration || typeof configuration !== "object") throw new TypeError("Grouped XLSX options require host-authored labels.");
  const options = captureFields(configuration, [...workbookOptionFields, "labelColumn", "subtotalLabel", "grandTotalLabel", "includeGrandTotal", "collapseGroups", "sheet", "groupStyle", "subtotalStyle", "grandTotalStyle"] as readonly (keyof GroupedXlsxOptions)[]);
  const sheet = captureFields(options.sheet ?? {}, [...sheetOptionFields, "name"] as readonly (keyof NonNullable<GroupedXlsxOptions["sheet"]>)[]);
  if (!options.labelColumn || typeof options.labelColumn !== "object") throw new TypeError("Grouped XLSX needs a leading label column.");
  const labelColumn = captureFields(options.labelColumn, ["header", "width", "wrapText", "alignment"]);
  if (typeof labelColumn.header !== "string") throw new TypeError("The leading label column needs a string header.");
  for (const label of [options.subtotalLabel, options.grandTotalLabel]) if (typeof label !== "function") throw new TypeError("Grouped summary labels must be functions.");
  for (const flag of [options.includeGrandTotal, options.collapseGroups]) if (flag !== undefined && typeof flag !== "boolean") throw new TypeError("Grouped XLSX flags must be booleans.");
  for (const key of unsupported) if ((sheet as Record<string, unknown>)[key] !== undefined) throw new TypeError("Grouped XLSX owns worksheet " + key + ".");
  if ((sheet as { autoFilter?: unknown }).autoFilter !== undefined && (sheet as { autoFilter?: unknown }).autoFilter !== false) throw new TypeError("Grouped summary rows cannot use a native table or autofilter.");
  if ((options as { sink?: unknown }).sink !== undefined || (options as { cellValueWriters?: unknown }).cellValueWriters !== undefined) throw new TypeError("Grouped XLSX owns its destination and uses already resolved source values.");
  for (const style of [options.groupStyle, options.subtotalStyle, options.grandTotalStyle]) if (style !== undefined) portableSheet({ alternatingRowStyle: style });
  for (const column of report.columns) if (Object.prototype.hasOwnProperty.call(report.totals, column.key!) && column.type !== "number") throw new TypeError("Editable grouped totals require explicitly numeric columns.");
  return { ...options, labelColumn, sheet: portableSheet(sheet),
    ...(options.groupStyle ? { groupStyle: copyStylePatch(options.groupStyle) as PortableCellStyle } : {}),
    ...(options.subtotalStyle ? { subtotalStyle: copyStylePatch(options.subtotalStyle) as PortableCellStyle } : {}),
    ...(options.grandTotalStyle ? { grandTotalStyle: copyStylePatch(options.grandTotalStyle) as PortableCellStyle } : {}) };
}

function label(value: unknown): string {
  if (typeof (value as { then?: unknown } | null)?.then === "function") void Promise.resolve(value).catch(() => {});
  if (typeof value !== "string") throw new TypeError("Grouped summary labels must return strings.");
  return value;
}

/** SUBTOTAL ignores nested subtotal formulas; functions 1..11 include collapsed outline details. */
function subtotal(operation: keyof typeof subtotalOperations, letter: string, first: number, last: number, cached: number | null): FormulaCell {
  const range = letter + first + ":" + letter + last;
  const expression = "SUBTOTAL(" + subtotalOperations[operation] + "," + range + ")";
  return new FormulaCell(operation === "average" || operation === "min" || operation === "max"
    ? "IF(SUBTOTAL(2," + range + ")=0,\"\"," + expression + ")" : expression, cached, operation === "count" ? 0 : undefined);
}

async function write(report: GroupedReportSource, options: GroupedXlsxOptions, sink?: import("../core/sinks.js").ByteSink): Promise<{ blob?: Blob; result: ExportResult }> {
  const { labelColumn, subtotalLabel, grandTotalLabel, includeGrandTotal, collapseGroups, sheet: sheetOptions, groupStyle, subtotalStyle, grandTotalStyle, onProgress, ...settings } = options;
  const operation = groupedReportOperation(report, settings.signal);
  let rows = 0, currentKind: GroupedReportEvent["kind"] = "groupStart";
  const progress = onProgress ? (event: ExportProgress) => onProgress({ ...event, rows, ...(event.sheetRows === undefined ? {} : { sheetRows: rows }) }) : undefined;
  let book: Workbook | undefined;
  try {
    book = new Workbook({ ...settings, ...(operation.signal ? { signal: operation.signal } : {}), ...(progress ? { onProgress: progress } : {}), ...(sink ? { sink } : {}) });
    const { name = "Data", ...sheetSettings } = sheetOptions ?? {};
    const columns = [{ ...labelColumn, type: "string" }, ...report.columns] as readonly XlsxColumn[];
    const styles = { groupStart: groupStyle ?? { font: { bold: true } }, groupEnd: subtotalStyle ?? { font: { bold: true } }, complete: grandTotalStyle ?? { font: { bold: true } } };
    const sheet = book.addWorksheet(name, { boldHeader: true, autoSize: { minWidth: 6, maxWidth: 54 }, ...sheetSettings, columns,
      maxOutlineLevel: 1,
      rowStyle: () => currentKind === "row" ? undefined : styles[currentKind],
      rowOutline: () => currentKind === "row" ? { level: 1, hidden: collapseGroups ?? false } : currentKind === "groupEnd" && collapseGroups ? { collapsed: true } : undefined });
    const operations = report.columns.map(column => Object.prototype.hasOwnProperty.call(report.totals, column.key!) ? report.totals[column.key!] : undefined);
    let firstDetail = 0, firstAnyDetail = 0, lastDetail = 0;
    async function* values(): AsyncGenerator<readonly (ExportValue | FormulaCell)[]> {
      for await (const event of operation.events) {
        currentKind = event.kind;
        if (event.kind === "groupStart") {
          firstDetail = sheet.rowCount + sheet.headerRowCount + 2;
          yield [event.label, ...report.columns.map(() => null)];
        } else if (event.kind === "row") {
          rows++;
          lastDetail = sheet.rowCount + sheet.headerRowCount + 1;
          firstAnyDetail ||= lastDetail;
          yield [null, ...event.values];
        } else if (event.kind === "groupEnd") {
          yield [label(subtotalLabel(event)), ...operations.map((operation, index) => operation
            ? subtotal(operation, columnName(index + 2), firstDetail, lastDetail, event.totals[index]!) : null)];
        } else if (includeGrandTotal !== false) {
          // Empty exports still retain a labelled numeric sum/count and blank average/min/max.
          yield [label(grandTotalLabel(event)), ...operations.map((operation, index) => operation
            ? firstAnyDetail ? subtotal(operation, columnName(index + 2), firstAnyDetail, lastDetail, event.totals[index]!) : event.totals[index]! : null)];
        }
      }
    }
    await sheet.addRows(values());
    if (sink) {
      const completion = await book.finish();
      return { result: { rows, columns: report.columns.length + 1, bytes: completion.bytes } };
    }
    const blob = await book.toBlob();
    return { blob, result: { rows, columns: report.columns.length + 1, bytes: blob.size } };
  } catch (error) { await book?.discard(error); throw error; }
  finally { operation.dispose(); }
}

/** Write one ordered grouped source to editable Excel outlines and finite native SUBTOTAL formulas. */
export async function writeGroupedXlsx(report: GroupedReportSource, options: GroupedXlsxOptions): Promise<Blob> {
  return (await write(report, prepare(report, options))).blob!;
}

/** Await accepted bytes; the caller retains its sink, partial bytes and stream lifetime. Counts exclude generated summary rows. */
export async function writeGroupedXlsxTo(report: GroupedReportSource, destination: OutputDestination, options: GroupedXlsxOptions): Promise<ExportResult> {
  const prepared = prepare(report, options);
  return withDestination(destination, async sink => (await write(report, prepared, sink)).result);
}

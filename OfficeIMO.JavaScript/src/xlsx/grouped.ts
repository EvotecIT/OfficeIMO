import type { ExportProgress, ExportResult, OutputDestination } from "../core/index.js";
import type { GroupedReportEvent, GroupedReportSource, GroupedReportSummary } from "../core/grouped-report.js";
import { groupedReportOperation } from "../core/grouped-report.js";
import { checkAbort, withAbort } from "../core/iteration.js";
import { withDestination } from "../core/sinks.js";
import type { ExportValue } from "../core/presentation.js";
import { captureFields } from "../internal/capture.js";
import { Workbook } from "./workbook.js";
import { FormulaCell } from "./formula.js";
import { columnName } from "./values.js";
import type { WorkbookOptions, XlsxColumn } from "./types.js";
import { captureWorkbookLimits, workbookOptionFields } from "./options.js";
import { retainDataRowRanges } from "./data-row-ranges.js";
import { groupedOptionFields, prepareGroupedSheet } from "./grouped-layout.js";
import type { GroupedXlsxSheetOptions, GroupedXlsxWorksheet, GroupedXlsxGroupRows } from "./grouped-layout.js";

export type { GroupedXlsxSheetOptions, GroupedXlsxWorksheet, GroupedXlsxGroupRows } from "./grouped-layout.js";

/** Convenience-writer settings plus the same composable analysis worksheet contract. */
export interface GroupedXlsxOptions extends Omit<WorkbookOptions, "sink" | "cellValueWriters">, GroupedXlsxSheetOptions {}

const subtotalOperations = { average: 1, count: 2, max: 4, min: 5, sum: 9 } as const;
type MutableGroupRows = { -readonly [Key in keyof GroupedXlsxGroupRows]: GroupedXlsxGroupRows[Key] };

function label(value: unknown): string {
  if (typeof (value as { then?: unknown } | null)?.then === "function") void Promise.resolve(value).catch(() => {});
  if (typeof value !== "string") throw new TypeError("Grouped summary labels must return strings.");
  return value;
}

/** SUBTOTAL ignores nested subtotal expressions; functions 1..11 include collapsed outline details. */
function subtotal(operation: keyof typeof subtotalOperations, letter: string, first: number, last: number, cached: number | null): FormulaCell {
  const range = letter + first + ":" + letter + last;
  const expression = "SUBTOTAL(" + subtotalOperations[operation] + "," + range + ")";
  return new FormulaCell(operation === "average" || operation === "min" || operation === "max"
    ? "IF(SUBTOTAL(2," + range + ")=0,\"\"," + expression + ")" : expression, cached, operation === "count" ? 0 : undefined);
}

async function append(book: Workbook, report: GroupedReportSource, options: GroupedXlsxSheetOptions,
  operation: ReturnType<typeof groupedReportOperation>, accepted?: (rows: number) => void): Promise<GroupedXlsxWorksheet> {
  checkAbort(operation.signal);
  const { labelColumn, subtotalLabel, grandTotalLabel, includeGrandTotal, collapseGroups, sheet: sheetOptions, groupStyle, subtotalStyle, grandTotalStyle, conditionalFormats, dataValidation } = options;
  let rows = 0, current: GroupedReportEvent | undefined;
  const groups: MutableGroupRows[] = [], stack: MutableGroupRows[] = [], detailRanges: { firstRow: number; lastRow: number }[] = [];
  let firstDetail = 0, lastDetail = 0, grandTotalRow: number | undefined;
  let grandSummary: GroupedReportSummary | undefined;
  const { name = "Data", ...sheetSettings } = sheetOptions ?? {};
  const columns = [{ ...labelColumn, type: "string" }, ...report.columns] as readonly XlsxColumn[];
  const styles = { groupStart: groupStyle ?? { font: { bold: true } }, groupEnd: subtotalStyle ?? { font: { bold: true } }, complete: grandTotalStyle ?? { font: { bold: true } } };
  const sheet = book.addWorksheet(name, { boldHeader: true, autoSize: { minWidth: 6, maxWidth: 54 }, ...sheetSettings, columns,
    ...(conditionalFormats === undefined ? {} : { conditionalFormats }), ...(dataValidation === undefined ? {} : { dataValidation }),
    maxOutlineLevel: report.groupBy.length,
    rowStyle: () => current && current.kind !== "row" ? styles[current.kind] : undefined,
    rowOutline: () => {
      if (!current || current.kind === "complete") return undefined;
      const level = current.kind === "row" ? current.path.length : current.level;
      return { level, hidden: !!collapseGroups && level > 0, collapsed: !!collapseGroups && current.kind === "groupEnd" };
    } });
  retainDataRowRanges(sheet, detailRanges);
  const operations = report.columns.map(column => Object.prototype.hasOwnProperty.call(report.totals, column.key!) ? report.totals[column.key!] : undefined);
  async function* values(): AsyncGenerator<readonly (ExportValue | FormulaCell)[]> {
    for await (const event of operation.events) {
      current = event;
      const worksheetRow = sheet.rowCount + sheet.headerRowCount + 1;
      if (event.kind === "groupStart") {
        const group: MutableGroupRows = { groupIndex: event.groupIndex, level: event.level, key: event.key, label: event.label, path: event.path,
          rows: 0, totals: Object.freeze([]), startRow: worksheetRow, firstDetailRow: 0, lastDetailRow: 0, summaryRow: 0 };
        groups.push(group); stack.push(group);
        yield [event.label, ...report.columns.map(() => null)];
      } else if (event.kind === "row") {
        rows++; accepted?.(rows); firstDetail ||= worksheetRow; lastDetail = worksheetRow;
        const previous = detailRanges.at(-1);
        if (previous && previous.lastRow + 1 === worksheetRow) previous.lastRow = worksheetRow;
        else detailRanges.push({ firstRow: worksheetRow, lastRow: worksheetRow });
        for (const group of stack) { group.firstDetailRow ||= worksheetRow; group.lastDetailRow = worksheetRow; }
        yield [null, ...event.values];
      } else if (event.kind === "groupEnd") {
        const group = stack.pop()!; group.summaryRow = worksheetRow; group.rows = event.rows; group.totals = event.totals;
        yield [label(subtotalLabel(event)), ...operations.map((operation, index) => operation
          ? subtotal(operation, columnName(index + 2), group.firstDetailRow, group.lastDetailRow, event.totals[index]!) : null)];
      } else {
        grandSummary = Object.freeze({ rows: event.rows, groups: event.groups, totals: event.totals });
        if (includeGrandTotal !== false) {
          grandTotalRow = worksheetRow;
          // An empty source still retains numeric sum/count and blank average/min/max.
          yield [label(grandTotalLabel(event)), ...operations.map((operation, index) => operation
            ? firstDetail ? subtotal(operation, columnName(index + 2), firstDetail, lastDetail, event.totals[index]!) : event.totals[index]! : null)];
        }
      }
    }
  }
  try {
    checkAbort(operation.signal);
    await withAbort(sheet.addRows(values()), operation.signal);
    // Grouped coordinates and native editing ranges cannot change after the result is returned.
    await withAbort(sheet.close(), operation.signal);
    return Object.freeze({ worksheet: sheet, rows, columns: columns.length,
      detailRanges: Object.freeze(detailRanges.map(range => Object.freeze({ ...range }))),
      groups: Object.freeze(groups.map(group => Object.freeze({ ...group }))), grandSummary: grandSummary!,
      ...(grandTotalRow === undefined ? {} : { grandTotalRow }) });
  } catch (error) { await book.discard(error); throw error; }
}

/** Add one ordered analysis worksheet, preserving the workbook's styles, budgets, destination and sibling sheets.
 * The returned worksheet is closed; finalize the workbook only after adding any remaining sheets. */
export async function addGroupedXlsx(book: Workbook, report: GroupedReportSource, options: GroupedXlsxSheetOptions): Promise<GroupedXlsxWorksheet> {
  const prepared = prepareGroupedSheet(report, options, book.settings.limits?.maxConditionalFormats ?? 1000, book.availableDataValidations), operation = groupedReportOperation(report, book.settings.signal);
  try { return await append(book, report, prepared, operation); }
  finally { operation.dispose(); }
}

function prepare(report: GroupedReportSource, configuration: GroupedXlsxOptions): GroupedXlsxOptions {
  if (!configuration || typeof configuration !== "object") throw new TypeError("Grouped XLSX options require host-authored labels.");
  const options = captureFields(configuration, [...workbookOptionFields, ...groupedOptionFields] as readonly (keyof GroupedXlsxOptions)[]);
  if ((options as { sink?: unknown }).sink !== undefined || (options as { cellValueWriters?: unknown }).cellValueWriters !== undefined)
    throw new TypeError("Grouped XLSX owns its destination and uses already resolved source values.");
  const limits = options.limits === undefined ? undefined : captureWorkbookLimits(options.limits);
  return { ...options, ...(limits === undefined ? {} : { limits }),
    ...prepareGroupedSheet(report, options, limits?.maxConditionalFormats ?? 1000, limits?.maxDataValidations ?? 1000) };
}

async function write(report: GroupedReportSource, options: GroupedXlsxOptions, sink?: import("../core/sinks.js").ByteSink): Promise<{ blob?: Blob; result: ExportResult }> {
  const { labelColumn: _labelColumn, subtotalLabel: _subtotalLabel, grandTotalLabel: _grandTotalLabel, includeGrandTotal: _includeGrandTotal,
    collapseGroups: _collapseGroups, sheet: _sheet, groupStyle: _groupStyle, subtotalStyle: _subtotalStyle, grandTotalStyle: _grandTotalStyle,
    conditionalFormats: _conditionalFormats, dataValidation: _dataValidation, onProgress, ...settings } = options;
  const operation = groupedReportOperation(report, settings.signal);
  let rows = 0, book: Workbook | undefined;
  const progress = onProgress ? (event: ExportProgress) => onProgress({ ...event, rows, ...(event.sheetRows === undefined ? {} : { sheetRows: rows }) }) : undefined;
  try {
    book = new Workbook({ ...settings, ...(operation.signal ? { signal: operation.signal } : {}), ...(progress ? { onProgress: progress } : {}), ...(sink ? { sink } : {}) });
    const result = await append(book, report, options, operation, accepted => { rows = accepted; });
    if (sink) {
      const completion = await book.finish();
      return { result: { rows, columns: result.columns, bytes: completion.bytes } };
    }
    const blob = await book.toBlob();
    return { blob, result: { rows, columns: result.columns, bytes: blob.size } };
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

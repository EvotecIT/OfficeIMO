import type { ColumnSettings } from "../core/index.js";
import type { GroupedReportEvent, GroupedReportSource, GroupedReportGroup, GroupedReportSummary } from "../core/grouped-report.js";
import { assertGroupedReportSource } from "../core/grouped-report.js";
import { captureFields } from "../internal/capture.js";
import { OfficeIMOError } from "../core/errors.js";
import { fields } from "./differential-styles.js";
import { sheetOptionFields } from "./options.js";
import { portableSheet } from "./portable.js";
import type { PortableCellStyle, PortableSheetOptions } from "./portable.js";
import { resolveDataColumn } from "./ranges.js";
import { copyStylePatch } from "./styles.js";
import type { ConditionalFormat } from "./conditional-types.js";
import type { ListValidation } from "./validation-types.js";
import type { Worksheet } from "./worksheet.js";
import type { DataRowRange } from "./data-row-ranges.js";

type GroupEnd = Extract<GroupedReportEvent, { readonly kind: "groupEnd" }>;
type Complete = Extract<GroupedReportEvent, { readonly kind: "complete" }>;
type GroupedSheet = Omit<PortableSheetOptions, "columns" | "headerRows" | "footer" | "table" | "autoFilter" | "mergedCells" | "hyperlinks" | "conditionalFormats" | "dataValidation" | "rowStyle" | "cellStyle" | "rowOutline" | "maxOutlineLevel">;

/** Layout for one grouped analysis worksheet in a caller-owned workbook. */
export interface GroupedXlsxSheetOptions {
  /** A separate label column preserves every visible typed data column. */
  readonly labelColumn: Pick<ColumnSettings, "header" | "width" | "wrapText" | "alignment">;
  readonly subtotalLabel: (group: GroupEnd) => string;
  readonly grandTotalLabel: (summary: Complete) => string;
  readonly includeGrandTotal?: boolean;
  /** Hidden outline details remain included in the native SUBTOTAL calculations. */
  readonly collapseGroups?: boolean;
  readonly sheet?: GroupedSheet & { readonly name?: string };
  readonly groupStyle?: PortableCellStyle;
  readonly subtotalStyle?: PortableCellStyle;
  readonly grandTotalStyle?: PortableCellStyle;
  /** Data-only ranges use stable data-column keys or one-based visible data-column numbers. A1 ranges and the generated label column are excluded. */
  readonly conditionalFormats?: readonly ConditionalFormat[];
  /** Inline editing lists apply only to detail rows, using the same data-column selectors. */
  readonly dataValidation?: readonly ListValidation[];
}

/** Stable one-based physical coordinates for one generated group and its body. */
export interface GroupedXlsxGroupRows extends GroupedReportGroup {
  readonly path: readonly GroupedReportGroup[];
  readonly rows: number;
  /** The shared source's column-order aggregates, without reaggregating worksheet cells. */
  readonly totals: readonly (number | null)[];
  readonly startRow: number;
  readonly firstDetailRow: number;
  readonly lastDetailRow: number;
  readonly summaryRow: number;
}

/** A completed analysis worksheet; the owning workbook remains open for sibling sheets and finalization. */
export interface GroupedXlsxWorksheet {
  readonly worksheet: Worksheet;
  /** Accepted source details, excluding generated headings and summaries. */
  readonly rows: number;
  /** Includes the generated label column. */
  readonly columns: number;
  /** Ordered physical detail runs, excluding all generated headings and summaries. Their count is bounded by the source's maxGroups. */
  readonly detailRanges: readonly DataRowRange[];
  /** Opening order, bounded by the source's maxGroups across all hierarchy levels. */
  readonly groups: readonly GroupedXlsxGroupRows[];
  /** The same one-pass source's complete aggregates, including when the grand-total row is omitted. */
  readonly grandSummary: GroupedReportSummary;
  readonly grandTotalRow?: number;
}

/** @internal */
export const groupedOptionFields = ["labelColumn", "subtotalLabel", "grandTotalLabel", "includeGrandTotal", "collapseGroups", "sheet", "groupStyle", "subtotalStyle", "grandTotalStyle", "conditionalFormats", "dataValidation"] as const;
const unsupported = ["columns", "headerRows", "footer", "table", "mergedCells", "hyperlinks", "conditionalFormats", "dataValidation", "rowStyle", "cellStyle", "rowOutline", "maxOutlineLevel"] as const;

function detailColumn(report: GroupedReportSource, selector: unknown): number {
  return resolveDataColumn(report.columns, selector as string | number, "Grouped detail") + 1;
}

function conditionalRules(report: GroupedReportSource, rules: readonly ConditionalFormat[] | undefined, maximum: number): readonly ConditionalFormat[] | undefined {
  if (rules === undefined) return undefined;
  if (!Array.isArray(rules)) throw new TypeError("conditionalFormats must be an array.");
  if (rules.length > maximum) throw new OfficeIMOError("RESOURCE_LIMIT", "maxConditionalFormats exceeded.");
  return rules.map(input => {
    const rule = captureFields(input as unknown as Record<string, unknown>, ["type", "range", "style", "stopIfTrue", "operator", "value", "values", "formula", "stops", "color", "minimum", "maximum", "showValue"]);
    const inputRange = rule.range;
    if (typeof inputRange === "string") throw new TypeError("Grouped conditional formats require detail-column selectors rather than A1 or generated-row ranges.");
    fields(inputRange, ["column", "through"], "grouped conditional range");
    const range = captureFields(inputRange as Record<string, unknown>, ["column", "through"]);
    return { ...rule, range: { column: detailColumn(report, range.column), ...(range.through === undefined ? {} : { through: detailColumn(report, range.through) }) } } as unknown as ConditionalFormat;
  });
}

function validationRules(report: GroupedReportSource, rules: readonly ListValidation[] | undefined, maximum: number): readonly ListValidation[] | undefined {
  if (rules === undefined) return undefined;
  if (!Array.isArray(rules)) throw new TypeError("dataValidation must be an array.");
  if (rules.length > maximum) throw new OfficeIMOError("RESOURCE_LIMIT", "maxDataValidations exceeded.");
  return rules.map(input => {
    const rule = captureFields(input, ["type", "column", "items", "allowBlank", "showErrorMessage", "errorStyle", "errorTitle", "errorMessage"]);
    return { ...rule, column: detailColumn(report, rule.column) };
  });
}

/** @internal Capture options before creating a worksheet or touching a one-pass source. */
export function prepareGroupedSheet(report: GroupedReportSource, configuration: GroupedXlsxSheetOptions, maxConditionalFormats = 1000, maxDataValidations = 1000): GroupedXlsxSheetOptions {
  assertGroupedReportSource(report);
  if (!configuration || typeof configuration !== "object") throw new TypeError("Grouped XLSX options require host-authored labels.");
  const options = captureFields(configuration, groupedOptionFields);
  const sheet = captureFields(options.sheet ?? {}, [...sheetOptionFields, "name"] as readonly (keyof NonNullable<GroupedXlsxSheetOptions["sheet"]>)[]);
  if (!options.labelColumn || typeof options.labelColumn !== "object") throw new TypeError("Grouped XLSX needs a leading label column.");
  const labelColumn = captureFields(options.labelColumn, ["header", "width", "wrapText", "alignment"]);
  if (typeof labelColumn.header !== "string") throw new TypeError("The leading label column needs a string header.");
  for (const label of [options.subtotalLabel, options.grandTotalLabel]) if (typeof label !== "function") throw new TypeError("Grouped summary labels must be functions.");
  for (const flag of [options.includeGrandTotal, options.collapseGroups]) if (flag !== undefined && typeof flag !== "boolean") throw new TypeError("Grouped XLSX flags must be booleans.");
  for (const key of unsupported) if ((sheet as Record<string, unknown>)[key] !== undefined) throw new TypeError("Grouped XLSX owns worksheet " + key + ".");
  if ((sheet as { autoFilter?: unknown }).autoFilter !== undefined && (sheet as { autoFilter?: unknown }).autoFilter !== false)
    throw new TypeError("Grouped summary rows cannot use a native table or autofilter.");
  for (const style of [options.groupStyle, options.subtotalStyle, options.grandTotalStyle]) if (style !== undefined) portableSheet({ alternatingRowStyle: style });
  for (const column of report.columns) if (Object.prototype.hasOwnProperty.call(report.totals, column.key!) && column.type !== "number")
    throw new TypeError("Editable grouped totals require explicitly numeric columns.");
  const conditionalFormats = conditionalRules(report, options.conditionalFormats, maxConditionalFormats), dataValidation = validationRules(report, options.dataValidation, maxDataValidations);
  return { ...options, labelColumn, sheet: portableSheet(sheet),
    ...(conditionalFormats === undefined ? {} : { conditionalFormats }), ...(dataValidation === undefined ? {} : { dataValidation }),
    ...(options.groupStyle ? { groupStyle: copyStylePatch(options.groupStyle) as PortableCellStyle } : {}),
    ...(options.subtotalStyle ? { subtotalStyle: copyStylePatch(options.subtotalStyle) as PortableCellStyle } : {}),
    ...(options.grandTotalStyle ? { grandTotalStyle: copyStylePatch(options.grandTotalStyle) as PortableCellStyle } : {}) };
}

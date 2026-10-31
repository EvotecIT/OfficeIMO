import type { RowOutline, PortableSheetOptions } from "../../xlsx/index.js";
import type { CanopyCapture, CanopyExportColumn, CanopyExportRow, CanopyPivotAxisPresentation } from "./types.js";

/** One serial enumeration's current metadata; the worksheet serializes each yielded row before requesting the next. */
export interface CanopyPivotMapping {
  readonly columns: readonly CanopyExportColumn[];
  readonly sheet: PortableSheetOptions;
  observe(row: CanopyExportRow): void;
}

function presentation(input: Partial<CanopyPivotAxisPresentation>, axes: ReadonlySet<string>): RowOutline {
  const { axisKey, outlineLevel, summary, grandTotal, hidden, collapsed } = input;
  if (typeof axisKey !== "string" || !axes.has(axisKey) || !Number.isInteger(outlineLevel) || outlineLevel! < 0 || outlineLevel! > 7 ||
    [summary, grandTotal, hidden, collapsed].some(flag => typeof flag !== "boolean")) throw new TypeError("Invalid frozen Canopy pivot axis presentation.");
  return Object.freeze({ level: outlineLevel!, hidden: hidden!, collapsed: collapsed! });
}

/** Qualify the explicit pivot path before flat/grouped handling; never evaluate query or aggregate data. */
export function canopyPivot(request: CanopyCapture["request"]): CanopyPivotMapping | undefined {
  if (request.pivot === undefined) return undefined;
  const pivot = request.pivot;
  if (!pivot || typeof pivot !== "object" || request.grouping !== undefined || !Array.isArray(request.columns) ||
    request.scope !== "filtered" || !Number.isSafeInteger(request.recordCount) || request.recordCount! < 0 ||
    !pivot.definition || !Array.isArray(pivot.definition.rows) || pivot.definition.rows.length > 4 ||
    !Array.isArray(pivot.definition.columns) || pivot.definition.columns.length > 2 ||
    !Array.isArray(pivot.definition.values) || !pivot.definition.values.length || pivot.definition.values.length > 32 ||
    pivot.definition.values.some(value => typeof value !== "string") ||
    !Number.isSafeInteger(pivot.distinctColumns) || pivot.distinctColumns < 0 || typeof pivot.columnsLimited !== "boolean" ||
    !Array.isArray(pivot.retainedColumns) || !Array.isArray(pivot.columnGroups)) throw new TypeError("Invalid explicit Canopy pivot descriptor.");
  const axes = (nodes: typeof pivot.rowAxis): ReadonlySet<string> => {
    if (!Array.isArray(nodes) || nodes.length > 100001) throw new TypeError("Pivot axes require bounded accepted identities.");
    const keys = new Set<string>();
    for (const node of nodes) {
      if (!node || typeof node.key !== "string" || !node.key || keys.has(node.key) || !Array.isArray(node.path) || typeof node.other !== "boolean")
        throw new TypeError("Invalid or duplicate Canopy pivot axis identity.");
      keys.add(node.key);
    }
    for (const node of nodes) if (node.parent !== undefined && !keys.has(node.parent)) throw new TypeError("Unknown Canopy pivot parent identity.");
    return keys;
  };
  const rowKeys = axes(pivot.rowAxis), columnKeys = axes(pivot.columnAxis);
  const columns = request.columns.map((column, index) => {
    if (index === 0 && column.axisKey === undefined && column.kind === "text") return Object.freeze({ ...column });
    const outline = presentation(column, columnKeys);
    return Object.freeze({ ...column, outlineLevel: outline.level!, hidden: outline.hidden!, collapsed: outline.collapsed! });
  });
  // Ranges are descriptor metadata, not another body. Validate them against the same ordered columns.
  if (pivot.columnGroups.length > columns.length) throw new RangeError("Pivot column group count exceeds projected columns.");
  for (const group of pivot.columnGroups) {
    if (!group || !columnKeys.has(group.key) || !Number.isSafeInteger(group.firstColumn) || group.firstColumn < 0 ||
      !Number.isSafeInteger(group.columnCount) || group.columnCount < 1 || group.firstColumn + group.columnCount > columns.length ||
      !Number.isInteger(group.depth) || group.depth < 1 || group.depth > 7 ||
      typeof group.hidden !== "boolean" || typeof group.collapsed !== "boolean") throw new TypeError("Invalid Canopy pivot column range.");
  }
  let current: RowOutline | undefined;
  const sheet: PortableSheetOptions = Object.freeze({ outlineSummaryBelow: false, outlineSummaryRight: false,
    maxOutlineLevel: 7, rowOutline: () => current,
    columnOutlines: Object.freeze(columns.flatMap((column, firstColumn) => column.axisKey === undefined ? [] :
      [Object.freeze({ firstColumn, columnCount: 1, level: column.outlineLevel!, hidden: column.hidden!, collapsed: column.collapsed! })])) });
  return Object.freeze({ columns: Object.freeze(columns), sheet,
    observe(row: CanopyExportRow) { current = presentation(row, rowKeys); } });
}

/** Per-cell formats preserve already-computed fractional percentages, ranks and differences. */
export function pivotNumberFormat(cell: { readonly mode?: string; readonly numberFormat?: string }): string | undefined {
  // Raw-value aggregates need no decimal declaration. Transformed values still
  // require the producer's frozen format (especially percent scaling).
  if (cell.mode === "value" && cell.numberFormat === undefined) return undefined;
  if (!["value", "percentGrandTotal", "percentParent", "percentColumn", "percentRow", "runningTotal", "rank", "difference", "percentChange"].includes(cell.mode ?? "") ||
    typeof cell.numberFormat !== "string" || !cell.numberFormat.length) throw new TypeError("Numeric pivot cells require a frozen mode and native numberFormat.");
  return cell.numberFormat;
}

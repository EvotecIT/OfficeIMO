import { ExportCell, assertScalar } from "../core/presentation.js";
import type { ExportValue } from "../core/presentation.js";
import type { CellStyle, Font, Fill, Border } from "./styles.js";
import type { WorkbookOptions, SheetOptions, CellWriterContext, RowStyleContext, CellStyleContext, ReportTitle, FooterOptions } from "./types.js";

/** Style definitions that do not refer to a workbook's component indexes. */
export interface PortableCellStyle extends Omit<CellStyle, "font" | "fill" | "border" | "numberFormat"> {
  readonly font?: Font; readonly fill?: Fill; readonly border?: Border; readonly numberFormat?: string;
}
/** Settings for writers that own their workbook and accept resolved portable values. */
export interface PortableWorkbookOptions extends Omit<WorkbookOptions, "sink" | "cellValueWriters"> {
  readonly cellValueWriters?: Readonly<Record<string, (value: import("../core/index.js").CellValue, context: CellWriterContext) => ExportValue>>;
}
/** Report layout and presentation without workbook-local style or Cell indexes. */
export interface PortableSheetOptions extends Omit<SheetOptions, "columns" | "headerStyle" | "alternatingRowStyle" | "rowStyle" | "cellStyle" | "title" | "footer"> {
  readonly alternatingRowStyle?: PortableCellStyle;
  readonly rowStyle?: (context: RowStyleContext) => PortableCellStyle | undefined;
  readonly cellStyle?: (context: CellStyleContext) => PortableCellStyle | undefined;
  readonly title?: Omit<ReportTitle, "style"> & { readonly style?: PortableCellStyle };
  readonly footer?: Omit<FooterOptions, "values" | "style"> & { readonly values?: readonly ExportValue[]; readonly style?: PortableCellStyle };
}
function style<T extends CellStyle | undefined>(patch: T): T {
  for (const component of ["font", "fill", "border", "numberFormat"] as const)
    if (typeof patch?.[component] === "number") throw new TypeError("Workbook-local style indexes require the advanced Workbook API; use style definitions.");
  return patch;
}
function value<T extends ExportValue>(result: T): T {
  if (!(result instanceof ExportCell)) assertScalar(result);
  return result;
}
/** @internal Qualify the portable boundary before source/destination activity. */
export function portableSheet(options: PortableSheetOptions = {}): PortableSheetOptions {
  if ((options as { headerStyle?: unknown }).headerStyle !== undefined)
    throw new TypeError("Workbook-local header styles require the advanced Workbook API; use boldHeader and headerFill.");
  style(options.alternatingRowStyle); style(options.title?.style); style(options.footer?.style);
  options.footer?.values?.forEach(value);
  const rowStyle = options.rowStyle, cellStyle = options.cellStyle;
  for (const callback of [rowStyle, cellStyle]) if (callback !== undefined && typeof callback !== "function") throw new TypeError("Style callbacks must be functions.");
  return { ...options, ...(rowStyle ? { rowStyle: (context: RowStyleContext) => style(rowStyle(context)) } : {}),
    ...(cellStyle ? { cellStyle: (context: CellStyleContext) => style(cellStyle(context)) } : {}) };
}
/** @internal Keep custom writers inside the same portable value contract. */
export function portableWorkbook(options: PortableWorkbookOptions = {}): PortableWorkbookOptions {
  const writers = options.cellValueWriters;
  if (!writers) return options;
  const wrapped: Record<string, (value: import("../core/index.js").CellValue, context: CellWriterContext) => ExportValue> = Object.create(null);
  for (const [type, writer] of Object.entries(writers)) {
    if (typeof writer !== "function") throw new TypeError("Cell value writers must be functions.");
    wrapped[type] = (input, context) => value(writer(input, context));
  }
  return { ...options, cellValueWriters: wrapped };
}

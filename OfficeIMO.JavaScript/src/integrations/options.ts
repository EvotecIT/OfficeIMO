import { captureFields } from "../internal/capture.js";
import { ExportBudget } from "../core/limits.js";
import type { ExportLimits } from "../core/limits.js";
import { captureCsvOptions as captureCsv } from "../csv/options.js";
import type { CsvOptions } from "../csv/index.js";
import { capturePdfOptions as capturePdf, capturePdfLimits, pdfLimitFields } from "../pdf/settings.js";
import type { PdfOptions, PdfLimits } from "../pdf/index.js";
import { captureWorkbookLimits, workbookLimitFields, workbookOptionFields, sheetOptionFields } from "../xlsx/options.js";
import type { XlsxOptions, WorkbookLimits } from "../xlsx/index.js";

type CommonLimits = WorkbookLimits & PdfLimits;
const limitFields = [...new Set<keyof CommonLimits>([...workbookLimitFields, ...pdfLimitFields])];
const portableWorkbookFields = workbookOptionFields.filter((key): key is Exclude<typeof workbookOptionFields[number], "sink"> => key !== "sink");

/** Shared adapter budgets override only explicitly supplied numeric format ceilings. */
export function captureCommonLimits(input: CommonLimits | undefined): Readonly<CommonLimits> | undefined {
  return input === undefined ? undefined : new ExportBudget(captureFields(input, limitFields)).limits;
}
export function mergeAdapterLimits<T extends ExportLimits>(format: T | undefined, common: Readonly<CommonLimits> | undefined): T {
  const result = { ...format };
  for (const [key, value] of Object.entries(common ?? {}))
    if (value !== undefined) (result as Record<string, unknown>)[key] = value;
  return result as T;
}
/** Snapshot the format owner's fields before an adapter adds captured columns and stream policy. */
export function captureXlsxOptions(input: Omit<XlsxOptions, "columns"> = {}): Omit<XlsxOptions, "columns"> {
  const options = captureFields(input, [...portableWorkbookFields, "sheet"]);
  const sheet = options.sheet && captureFields(options.sheet, [...sheetOptionFields, "name"] as readonly (keyof NonNullable<typeof options.sheet>)[]);
  const footer = sheet?.footer && captureFields(sheet.footer, ["rows", "values", "totals", "style"]);
  return { ...options, ...(sheet === undefined ? {} : { sheet: { ...sheet, ...(footer === undefined ? {} : { footer }) } }),
    ...(options.limits === undefined ? {} : { limits: captureWorkbookLimits(options.limits) }) };
}
export function captureCsvOptions(input: Omit<CsvOptions, "columns"> = {}): Omit<CsvOptions, "columns"> {
  const options = captureCsv(input);
  return { ...options, ...(options.limits === undefined ? {} : { limits: new ExportBudget(options.limits).limits }) };
}
export function capturePdfOptions(input: Omit<PdfOptions, "columns"> = {}): Omit<PdfOptions, "columns"> {
  const options = capturePdf(input);
  return { ...options, ...(options.limits === undefined ? {} : { limits: capturePdfLimits(options.limits) }) };
}

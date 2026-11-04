import { captureFields } from "../internal/capture.js";
import type { WorkbookLimits, WorkbookOptions } from "./types.js";

/** @internal The workbook owner captures these settings before validation or adapter merging. */
export const workbookOptionFields = ["signal", "onProgress", "limits", "sink", "dateMode", "compression", "invalidCharacterPolicy", "appProperties", "cellValueWriters", "oversizedText", "creator", "title", "subject", "description", "created", "modified"] as const;

/** @internal Worksheet settings also used by the portable sheet owner. */
export const sheetOptionFields = ["columns", "includeHeader", "headerRows", "freezeHeader", "autoFilter", "boldHeader", "headerFill", "headerStyle", "alternatingRowStyle", "rowStyle", "cellStyle", "rowOutline", "maxOutlineLevel", "outlineSummaryBelow", "outlineSummaryRight", "columnOutlines", "rowHeight", "headerHeight", "defaultColumnWidth", "freezeColumns", "table", "autoSize", "footer", "print", "title", "mergedCells", "hyperlinks", "conditionalFormats", "dataValidation"] as const;

/** @internal Includes inherited format-specific ceilings before a wrapper combines stream settings. */
export const workbookLimitFields = ["maxRows", "maxCells", "maxTextCharacters", "maxOutputBytes", "maxStyles", "maxSheets", "maxOverflowCharacters", "maxHyperlinks", "maxImageBytes", "maxBufferedCells", "maxBufferedCharacters", "maxMergedRanges", "maxConditionalFormats", "maxDifferentialStyles", "maxCharts", "maxChartPoints", "maxChartTextCharacters", "maxFormulaCharacters", "maxDataValidations", "maxValidationTextCharacters"] as const;

/** @internal Capture the workbook's known ceilings without dropping supported prototype fields. */
export function captureWorkbookLimits(limits: WorkbookLimits): WorkbookLimits {
  return captureFields(limits, workbookLimitFields);
}

/** @internal One immutable-input capture for direct workbooks and thin portable writers. */
export function captureWorkbookOptions<T extends WorkbookOptions>(options: T): T {
  const captured = captureFields(options, workbookOptionFields);
  return { ...captured, ...(captured.limits === undefined ? {} : { limits: captureWorkbookLimits(captured.limits) }) };
}

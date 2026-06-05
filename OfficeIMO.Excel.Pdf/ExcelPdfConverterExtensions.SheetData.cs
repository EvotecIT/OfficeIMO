using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static ExcelSheet? GetWorkbookSheet(ExcelDocument document, string sheetName) {
            foreach (ExcelSheet sheet in document.Sheets) {
                if (string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase)) {
                    return sheet;
                }
            }

            return null;
        }

        private static bool HasHiddenRowsOrColumns(ExcelSheet? workbookSheet) {
            if (workbookSheet == null) {
                return false;
            }

            return workbookSheet.GetRowDefinitions().Any(row => row.Hidden) ||
                   workbookSheet.GetColumnDefinitions().Any(column => column.Hidden);
        }

        private static string GetExportRange(ExcelSheetReader sheet, ExcelSheet? workbookSheet, ExcelPdfSaveOptions options) {
            string? printArea = GetWorksheetPrintArea(workbookSheet, options);
            if (!string.IsNullOrWhiteSpace(printArea)) {
                if (ContainsMultiplePrintAreas(printArea!)) {
                    AddWarning(
                        options,
                        sheet.Name,
                        "WorksheetPrintArea",
                        "Multi-area worksheet print areas are not supported by the first-party PDF exporter; exporting the worksheet used range instead.");
                    return sheet.GetUsedRangeA1();
                }

                return NormalizeA1Range(printArea!);
            }

            return sheet.GetUsedRangeA1();
        }

        private static bool ContainsMultiplePrintAreas(string printArea) {
            bool inQuotedSheetName = false;
            for (int i = 0; i < printArea.Length; i++) {
                char current = printArea[i];
                if (current == '\'') {
                    if (inQuotedSheetName && i + 1 < printArea.Length && printArea[i + 1] == '\'') {
                        i++;
                    } else {
                        inQuotedSheetName = !inQuotedSheetName;
                    }
                } else if (current == ',' && !inQuotedSheetName) {
                    return true;
                }
            }

            return false;
        }

        private static bool HasWorksheetPrintArea(ExcelSheet? workbookSheet, ExcelPdfSaveOptions options) =>
            !string.IsNullOrWhiteSpace(GetWorksheetPrintArea(workbookSheet, options));

        private static string? GetWorksheetPrintArea(ExcelSheet? workbookSheet, ExcelPdfSaveOptions options) =>
            options.UseWorksheetPrintAreas && workbookSheet != null ? workbookSheet.GetPrintArea() : null;

        private static SheetExportData ReadSheetExportData(ExcelSheetReader sheet, ExcelSheet? workbookSheet, string exportRange, ExcelPdfSaveOptions options, PdfCore.PdfStandardFont defaultFontFamily) {
            string normalizedRange = NormalizeA1Range(exportRange);
            A1.TryParseRange(normalizedRange, out int rangeFirstRow, out int rangeFirstColumn, out _, out int rangeLastColumn);
            RangeExportData bodyRange = ReadRangeExportData(sheet, workbookSheet, normalizedRange, options, defaultFontFamily);
            object?[,] values = bodyRange.Values;
            ExcelCellStyleSnapshot?[,]? styles = bodyRange.Styles;
            ExcelHyperlinkSnapshot?[,]? hyperlinks = bodyRange.Hyperlinks;
            string?[,]? cellReferences = bodyRange.CellReferences;
            MergeLayoutData? mergedCells = bodyRange.MergedCells;
            ColumnLayoutData? columnWidths = bodyRange.ColumnWidths;
            RowLayoutData? rowHeights = bodyRange.RowHeights;
            int headerRows = options.HeaderRowCount;
            if (!options.UseWorksheetPrintTitleRows || workbookSheet == null) {
                return CreateSheetExportData(workbookSheet, values, styles, hyperlinks, cellReferences, mergedCells, columnWidths, rowHeights, headerRows, rangeFirstRow, options);
            }

            ExcelPrintTitles titles = workbookSheet.GetPrintTitles();
            if (!titles.HasRows) {
                return CreateSheetExportData(workbookSheet, values, styles, hyperlinks, cellReferences, mergedCells, columnWidths, rowHeights, headerRows, rangeFirstRow, options);
            }

            int firstTitleRow = titles.FirstRow!.Value;
            int lastTitleRow = titles.LastRow!.Value;
            if (firstTitleRow < rangeFirstRow) {
                int prependedLastTitleRow = Math.Min(lastTitleRow, rangeFirstRow - 1);
                string titleRange = ToA1Range(firstTitleRow, rangeFirstColumn, prependedLastTitleRow, rangeLastColumn);
                RangeExportData titleRangeData = ReadRangeExportData(sheet, workbookSheet, titleRange, options, defaultFontFamily);
                int prependedRowCount = titleRangeData.Values.GetLength(0);
                int bodyRowCount = values.GetLength(0);
                int columnCount = values.GetLength(1);
                object?[,] prependedValues = PrependRows(titleRangeData.Values, values);
                ExcelCellStyleSnapshot?[,]? prependedStyles = PrependRows(titleRangeData.Styles, styles, prependedRowCount, bodyRowCount, columnCount);
                ExcelHyperlinkSnapshot?[,]? prependedHyperlinks = PrependRows(titleRangeData.Hyperlinks, hyperlinks, prependedRowCount, bodyRowCount, columnCount);
                string?[,]? prependedCellReferences = PrependRows(titleRangeData.CellReferences, cellReferences, prependedRowCount, bodyRowCount, columnCount);
                MergeLayoutData? prependedMergedCells = PrependRows(titleRangeData.MergedCells, mergedCells, prependedRowCount, bodyRowCount, columnCount);
                RowLayoutData? prependedRowHeights = PrependRows(titleRangeData.RowHeights, rowHeights, prependedRowCount, bodyRowCount);
                int overlappingTitleRows = lastTitleRow >= rangeFirstRow
                    ? Math.Min(bodyRowCount, lastTitleRow - rangeFirstRow + 1)
                    : 0;
                return CreateSheetExportData(
                    workbookSheet,
                    prependedValues,
                    prependedStyles,
                    prependedHyperlinks,
                    prependedCellReferences,
                    prependedMergedCells,
                    columnWidths,
                    prependedRowHeights,
                    Math.Max(headerRows, prependedRowCount + overlappingTitleRows),
                    rangeFirstRow,
                    options);
            }

            if (firstTitleRow <= rangeFirstRow && lastTitleRow >= rangeFirstRow) {
                int titleRowsInsideRange = Math.Min(values.GetLength(0), lastTitleRow - rangeFirstRow + 1);
                headerRows = Math.Max(headerRows, titleRowsInsideRange);
            }

            return CreateSheetExportData(workbookSheet, values, styles, hyperlinks, cellReferences, mergedCells, columnWidths, rowHeights, headerRows, rangeFirstRow, options);
        }

        private static SheetExportData CreateSheetExportData(ExcelSheet? workbookSheet, object?[,] values, ExcelCellStyleSnapshot?[,]? styles, ExcelHyperlinkSnapshot?[,]? hyperlinks, string?[,]? cellReferences, MergeLayoutData? mergedCells, ColumnLayoutData? columnWidths, RowLayoutData? rowHeights, int headerRows, int firstBodyRowNumber, ExcelPdfSaveOptions options) {
            ConditionalFillData? conditionalFills = ReadConditionalFillData(
                workbookSheet,
                values,
                cellReferences,
                options.UseWorksheetCellStyles);

            return new SheetExportData(values, styles, hyperlinks, cellReferences, mergedCells, columnWidths, rowHeights, headerRows, firstBodyRowNumber, conditionalFills);
        }

        private static RangeExportData ReadRangeExportData(ExcelSheetReader sheet, ExcelSheet? workbookSheet, string normalizedRange, ExcelPdfSaveOptions options, PdfCore.PdfStandardFont defaultFontFamily) {
            object?[,] rawValues = sheet.ReadRange(normalizedRange);
            VisibilityLayoutData? visibility = ReadVisibilityLayoutData(
                workbookSheet,
                normalizedRange,
                rawValues.GetLength(0),
                rawValues.GetLength(1),
                options.RespectWorksheetHiddenRowsAndColumns);
            object?[,] values = FilterValues(rawValues, visibility);
            int rowCount = values.GetLength(0);
            int columnCount = values.GetLength(1);
            string?[,]? cellReferences = ReadCellReferenceData(
                normalizedRange,
                rowCount,
                columnCount,
                visibility);
            ExcelCellStyleSnapshot?[,]? styles = ReadCellStyleData(
                workbookSheet,
                normalizedRange,
                rowCount,
                columnCount,
                options.UseWorksheetCellStyles,
                defaultFontFamily,
                visibility);
            ExcelHyperlinkSnapshot?[,]? hyperlinks = ReadHyperlinkData(
                workbookSheet,
                normalizedRange,
                rowCount,
                columnCount,
                options.UseWorksheetHyperlinks,
                visibility);
            MergeLayoutData? mergedCells = ReadMergeLayoutData(
                workbookSheet,
                normalizedRange,
                rowCount,
                columnCount,
                options.UseWorksheetMergedCells,
                visibility);
            ColumnLayoutData? columnWidths = ReadColumnLayoutData(
                workbookSheet,
                normalizedRange,
                columnCount,
                options.UseWorksheetColumnWidths,
                visibility);
            RowLayoutData? rowHeights = ReadRowLayoutData(
                workbookSheet,
                normalizedRange,
                rowCount,
                options.UseWorksheetRowHeights,
                visibility);

            return new RangeExportData(values, styles, hyperlinks, cellReferences, mergedCells, columnWidths, rowHeights);
        }

    }
}

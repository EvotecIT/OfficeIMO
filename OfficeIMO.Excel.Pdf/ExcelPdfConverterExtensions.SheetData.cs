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

        private static bool HasHiddenRowsOrColumns(ExcelSheet? workbookSheet, System.Threading.CancellationToken cancellationToken) {
            if (workbookSheet == null) {
                return false;
            }

            return workbookSheet.GetRowDefinitions().Any(row => row.Hidden) ||
                   workbookSheet.GetColumnDefinitions(cancellationToken).Any(column => column.Hidden);
        }

        private static IReadOnlyList<string> GetExportRanges(ExcelSheetReader sheet, ExcelSheet? workbookSheet, ExcelPdfExportOptions options) {
            IReadOnlyList<string> areas = options.UseWorksheetPrintAreas && workbookSheet != null
                ? workbookSheet.GetPrintAreas()
                : Array.Empty<string>();
            if (areas.Count == 0) return new[] { ExtendUsedRangeToMergedCells(sheet.GetUsedRangeA1(), workbookSheet, options) };

            ExcelA1.TryParseRange(NormalizeA1Range(sheet.GetUsedRangeA1()), out _, out _, out int usedLastRow, out int usedLastColumn);
            var ranges = new List<string>(areas.Count);
            foreach (string area in areas) {
                options.CancellationToken.ThrowIfCancellationRequested();
                if (!workbookSheet!.TryParsePrintAreaReference(area, out ExcelReference? reference) || reference == null) {
                    throw new ArgumentException("Worksheet print area must be a valid local A1 reference: " + area, nameof(options));
                }
                int firstRow = reference.Kind == ExcelReferenceKind.WholeColumn ? 1 : Math.Min(reference.Start.Row, reference.End.Row);
                int lastRow = reference.Kind == ExcelReferenceKind.WholeColumn ? Math.Max(1, usedLastRow) : Math.Max(reference.Start.Row, reference.End.Row);
                int firstColumn = reference.Kind == ExcelReferenceKind.WholeRow ? 1 : Math.Min(reference.Start.Column, reference.End.Column);
                int lastColumn = reference.Kind == ExcelReferenceKind.WholeRow ? Math.Max(1, usedLastColumn) : Math.Max(reference.Start.Column, reference.End.Column);
                ranges.Add(ToA1Range(firstRow, firstColumn, lastRow, lastColumn));
            }
            return ranges;
        }

        /// <summary>
        /// Excel prints a merged cell whole when its anchor is inside the used range, even
        /// if the other cells of the merge hold no value. Grow the automatic range to the
        /// merge extents so a title merged across A1:F1 is not cut at the last value column.
        /// </summary>
        private static string ExtendUsedRangeToMergedCells(string usedRange, ExcelSheet? workbookSheet, ExcelPdfExportOptions options) {
            if (workbookSheet == null || !options.UseWorksheetMergedCells || IsBoundedWorksheetRead(options) ||
                !ExcelA1.TryParseRange(NormalizeA1Range(usedRange), out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                return usedRange;
            }

            int extendedLastRow = lastRow;
            int extendedLastColumn = lastColumn;
            foreach (ExcelMergedRangeInfo merged in workbookSheet.GetMergedRanges()) {
                if (merged.StartRow < firstRow || merged.StartRow > lastRow ||
                    merged.StartColumn < firstColumn || merged.StartColumn > lastColumn) {
                    continue;
                }

                extendedLastRow = Math.Max(extendedLastRow, merged.EndRow);
                extendedLastColumn = Math.Max(extendedLastColumn, merged.EndColumn);
            }

            return extendedLastRow == lastRow && extendedLastColumn == lastColumn
                ? usedRange
                : ToA1Range(firstRow, firstColumn, extendedLastRow, extendedLastColumn);
        }

        private static bool IsBoundedWorksheetRead(ExcelPdfExportOptions options) =>
            options.UseBoundedWorksheetRead && options.MaxRowsPerSheet.HasValue;

        private static bool HasWorksheetPrintArea(ExcelSheet? workbookSheet, ExcelPdfExportOptions options) =>
            !string.IsNullOrWhiteSpace(GetWorksheetPrintArea(workbookSheet, options));

        private static string? GetWorksheetPrintArea(ExcelSheet? workbookSheet, ExcelPdfExportOptions options) =>
            options.UseWorksheetPrintAreas && workbookSheet != null ? workbookSheet.GetPrintArea() : null;

        private static SheetExportData ReadSheetExportData(ExcelDocument document, ExcelSheetReader sheet, ExcelSheet? workbookSheet, string exportRange, ExcelPdfExportOptions options, PdfCore.PdfStandardFont defaultFontFamily) {
            string normalizedRange = NormalizeA1Range(exportRange);
            ExcelA1.TryParseRange(normalizedRange, out int rangeFirstRow, out int rangeFirstColumn, out _, out int rangeLastColumn);
            bool boundedRead = IsBoundedWorksheetRead(options);
            ExcelSheet? metadataSheet = boundedRead ? null : workbookSheet;
            IReadOnlyList<StructuredTableVisualData> structuredTables = ReadStructuredTableVisuals(document, sheet.Name, options);
            RangeExportData bodyRange = ReadRangeWithPrintTitleColumns(sheet, metadataSheet, workbookSheet, normalizedRange, options, defaultFontFamily, boundedRead);
            object?[,] values = bodyRange.Values;
            ExcelCellStyleInfo?[,]? styles = bodyRange.Styles;
            ExcelHyperlinkInfo?[,]? hyperlinks = bodyRange.Hyperlinks;
            string?[,]? cellReferences = bodyRange.CellReferences;
            MergeLayoutData? mergedCells = bodyRange.MergedCells;
            ColumnLayoutData? columnWidths = bodyRange.ColumnWidths;
            RowLayoutData? rowHeights = bodyRange.RowHeights;
            int headerRows = options.HeaderRowCount;
            if (!options.UseWorksheetPrintTitleRows || metadataSheet == null) {
                return CreateSheetExportData(metadataSheet, values, styles, hyperlinks, cellReferences, mergedCells, columnWidths, rowHeights, headerRows, rangeFirstRow, structuredTables, options);
            }

            ExcelPrintTitles titles = metadataSheet.GetPrintTitles();
            if (!titles.HasRows) {
                return CreateSheetExportData(metadataSheet, values, styles, hyperlinks, cellReferences, mergedCells, columnWidths, rowHeights, headerRows, rangeFirstRow, structuredTables, options);
            }

            int firstTitleRow = titles.FirstRow!.Value;
            int lastTitleRow = titles.LastRow!.Value;
            if (firstTitleRow < rangeFirstRow) {
                int prependedLastTitleRow = Math.Min(lastTitleRow, rangeFirstRow - 1);
                string titleRange = ToA1Range(firstTitleRow, rangeFirstColumn, prependedLastTitleRow, rangeLastColumn);
                RangeExportData titleRangeData = ReadRangeWithPrintTitleColumns(sheet, metadataSheet, workbookSheet, titleRange, options, defaultFontFamily, boundedRead);
                int prependedRowCount = titleRangeData.Values.GetLength(0);
                int bodyRowCount = values.GetLength(0);
                int columnCount = values.GetLength(1);
                object?[,] prependedValues = PrependRows(titleRangeData.Values, values);
                ExcelCellStyleInfo?[,]? prependedStyles = PrependRows(titleRangeData.Styles, styles, prependedRowCount, bodyRowCount, columnCount);
                ExcelHyperlinkInfo?[,]? prependedHyperlinks = PrependRows(titleRangeData.Hyperlinks, hyperlinks, prependedRowCount, bodyRowCount, columnCount);
                string?[,]? prependedCellReferences = PrependRows(titleRangeData.CellReferences, cellReferences, prependedRowCount, bodyRowCount, columnCount);
                MergeLayoutData? prependedMergedCells = PrependRows(titleRangeData.MergedCells, mergedCells, prependedRowCount, bodyRowCount, columnCount);
                RowLayoutData? prependedRowHeights = PrependRows(titleRangeData.RowHeights, rowHeights, prependedRowCount, bodyRowCount);
                int overlappingTitleRows = lastTitleRow >= rangeFirstRow
                    ? Math.Min(bodyRowCount, lastTitleRow - rangeFirstRow + 1)
                    : 0;
                return CreateSheetExportData(
                    metadataSheet,
                    prependedValues,
                    prependedStyles,
                    prependedHyperlinks,
                    prependedCellReferences,
                    prependedMergedCells,
                    columnWidths,
                    prependedRowHeights,
                    Math.Max(headerRows, prependedRowCount + overlappingTitleRows),
                    rangeFirstRow,
                    structuredTables,
                    options);
            }

            if (firstTitleRow <= rangeFirstRow && lastTitleRow >= rangeFirstRow) {
                int titleRowsInsideRange = Math.Min(values.GetLength(0), lastTitleRow - rangeFirstRow + 1);
                headerRows = Math.Max(headerRows, titleRowsInsideRange);
            }

            return CreateSheetExportData(metadataSheet, values, styles, hyperlinks, cellReferences, mergedCells, columnWidths, rowHeights, headerRows, rangeFirstRow, structuredTables, options);
        }

        private static SheetExportData CreateSheetExportData(ExcelSheet? workbookSheet, object?[,] values, ExcelCellStyleInfo?[,]? styles, ExcelHyperlinkInfo?[,]? hyperlinks, string?[,]? cellReferences, MergeLayoutData? mergedCells, ColumnLayoutData? columnWidths, RowLayoutData? rowHeights, int headerRows, int firstBodyRowNumber, IReadOnlyList<StructuredTableVisualData> structuredTables, ExcelPdfExportOptions options) {
            ConditionalFillData? conditionalFills = ReadConditionalFillData(
                workbookSheet,
                values,
                cellReferences,
                options.UseWorksheetCellStyles);

            IReadOnlyList<int> titleColumns = GetPrintTitleColumnIndexes(workbookSheet, cellReferences, options);
            IReadOnlyList<int> titleRows = GetPrintTitleRowIndexes(workbookSheet, cellReferences, options);
            return new SheetExportData(values, styles, hyperlinks, cellReferences, mergedCells, columnWidths, rowHeights, headerRows, firstBodyRowNumber, structuredTables, conditionalFills, titleColumns, titleRows);
        }

        /// <summary>
        /// The reader returns a formula's text when the workbook stores no cached result.
        /// Excel shows the calculated value instead, so evaluate supported formulas here
        /// without changing the workbook, and report the ones that cannot be calculated.
        /// </summary>
        private static void ApplyUncachedFormulaValues(ExcelSheet? metadataSheet, string normalizedRange, object?[,] values, ExcelPdfExportOptions options) {
            if (metadataSheet == null || values.Length == 0 ||
                !ExcelA1.TryParseRange(normalizedRange, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                return;
            }

            var (calculated, unsupported) = metadataSheet.EvaluateUncachedFormulaDisplayValues(firstRow, firstColumn, lastRow, lastColumn);
            foreach (KeyValuePair<(int Row, int Column), object?> entry in calculated) {
                int row = entry.Key.Row - firstRow;
                int column = entry.Key.Column - firstColumn;
                if (row >= 0 && column >= 0 && row < values.GetLength(0) && column < values.GetLength(1)) {
                    values[row, column] = entry.Value;
                }
            }

            if (unsupported.Count > 0) {
                AddWarning(
                    options,
                    metadataSheet.Name,
                    "formula-not-calculated",
                    $"{unsupported.Count.ToString(CultureInfo.InvariantCulture)} formula cell(s) have no cached result and use functions the built-in calculator does not support; they were exported blank. Recalculate and save the workbook in a spreadsheet application to include their values. First cell: {unsupported[0]}.",
                    OfficeConversionLossKind.Omission);
            }
        }

        private static RangeExportData ReadRangeExportData(ExcelSheetReader sheet, ExcelSheet? metadataSheet, ExcelSheet? styleSheet, string normalizedRange, ExcelPdfExportOptions options, PdfCore.PdfStandardFont defaultFontFamily, bool boundedRead) {
            object?[,] rawValues = sheet.ReadRange(normalizedRange);
            ApplyUncachedFormulaValues(metadataSheet, normalizedRange, rawValues, options);
            VisibilityLayoutData? visibility = ReadVisibilityLayoutData(
                metadataSheet,
                normalizedRange,
                rawValues.GetLength(0),
                rawValues.GetLength(1),
                options.RespectWorksheetHiddenRowsAndColumns,
                options.CancellationToken);
            object?[,] values = FilterValues(rawValues, visibility);
            int rowCount = values.GetLength(0);
            int columnCount = values.GetLength(1);
            string?[,]? cellReferences = ReadCellReferenceData(
                normalizedRange,
                rowCount,
                columnCount,
                visibility);
            ExcelCellStyleInfo?[,]? styles = ReadCellStyleData(
                styleSheet,
                sheet,
                normalizedRange,
                rowCount,
                columnCount,
                options.UseWorksheetCellStyles,
                defaultFontFamily,
                boundedRead,
                visibility);
            ExcelHyperlinkInfo?[,]? hyperlinks = ReadHyperlinkData(
                metadataSheet,
                normalizedRange,
                rowCount,
                columnCount,
                options.UseWorksheetHyperlinks,
                visibility);
            MergeLayoutData? mergedCells = ReadMergeLayoutData(
                metadataSheet,
                normalizedRange,
                rowCount,
                columnCount,
                options.UseWorksheetMergedCells,
                visibility);
            ColumnLayoutData? columnWidths = ReadColumnLayoutData(
                metadataSheet,
                normalizedRange,
                columnCount,
                options.UseWorksheetColumnWidths,
                options.CancellationToken,
                visibility);
            RowLayoutData? rowHeights = ReadRowLayoutData(
                metadataSheet,
                normalizedRange,
                rowCount,
                options.UseWorksheetRowHeights,
                visibility);

            return new RangeExportData(values, styles, hyperlinks, cellReferences, mergedCells, columnWidths, rowHeights);
        }

    }
}

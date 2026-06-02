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

        private static SheetExportData ReadSheetExportData(ExcelSheetReader sheet, ExcelSheet? workbookSheet, string exportRange, ExcelPdfSaveOptions options) {
            string normalizedRange = NormalizeA1Range(exportRange);
            A1.TryParseRange(normalizedRange, out int rangeFirstRow, out int rangeFirstColumn, out _, out int rangeLastColumn);
            RangeExportData bodyRange = ReadRangeExportData(sheet, workbookSheet, normalizedRange, options);
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
                RangeExportData titleRangeData = ReadRangeExportData(sheet, workbookSheet, titleRange, options);
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

        private static RangeExportData ReadRangeExportData(ExcelSheetReader sheet, ExcelSheet? workbookSheet, string normalizedRange, ExcelPdfSaveOptions options) {
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

        private static ConditionalFillData? ReadConditionalFillData(ExcelSheet? workbookSheet, object?[,] values, string?[,]? cellReferences, bool enabled) {
            if (!enabled || workbookSheet == null || cellReferences == null) {
                return null;
            }

            IReadOnlyList<ExcelConditionalFormattingInfo> rules = workbookSheet.GetConditionalFormattingRules();
            if (rules.Count == 0) {
                return null;
            }

            var fills = new Dictionary<(int Row, int Column), string>();
            var dataBars = new Dictionary<(int Row, int Column), ConditionalDataBarCell>();
            var icons = new Dictionary<(int Row, int Column), ConditionalIconCell>();
            foreach (ExcelConditionalFormattingInfo rule in rules
                .Where(rule => string.Equals(rule.Type, "ColorScale", StringComparison.OrdinalIgnoreCase) && rule.ColorScaleColors.Count >= 2)
                .OrderByDescending(rule => rule.Priority)) {
                if (!TryGetRgb(rule.ColorScaleColors[0], out byte startR, out byte startG, out byte startB) ||
                    !TryGetRgb(rule.ColorScaleColors[rule.ColorScaleColors.Count - 1], out byte endR, out byte endG, out byte endB)) {
                    continue;
                }

                var candidates = new List<(int Row, int Column, double Value)>();
                for (int row = 0; row < values.GetLength(0); row++) {
                    for (int column = 0; column < values.GetLength(1); column++) {
                        string? cellReference = cellReferences[row, column];
                        if (!string.IsNullOrWhiteSpace(cellReference) &&
                            IsCellReferenceInReferenceList(cellReference!, rule.Range) &&
                            TryGetConditionalNumericValue(values[row, column], out double numericValue)) {
                            candidates.Add((row, column, numericValue));
                        }
                    }
                }

                if (candidates.Count == 0) {
                    continue;
                }

                double min = candidates.Min(candidate => candidate.Value);
                double max = candidates.Max(candidate => candidate.Value);
                foreach (var candidate in candidates) {
                    double ratio = max <= min ? 0.5D : Math.Max(0D, Math.Min(1D, (candidate.Value - min) / (max - min)));
                    fills[(candidate.Row, candidate.Column)] = InterpolateRgbHex(startR, startG, startB, endR, endG, endB, ratio);
                }
            }

            foreach (ExcelConditionalFormattingInfo rule in rules
                .Where(rule => string.Equals(rule.Type, "DataBar", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(rule.DataBarColor))
                .OrderByDescending(rule => rule.Priority)) {
                var candidates = new List<(int Row, int Column, double Value)>();
                for (int row = 0; row < values.GetLength(0); row++) {
                    for (int column = 0; column < values.GetLength(1); column++) {
                        string? cellReference = cellReferences[row, column];
                        if (!string.IsNullOrWhiteSpace(cellReference) &&
                            IsCellReferenceInReferenceList(cellReference!, rule.Range) &&
                            TryGetConditionalNumericValue(values[row, column], out double numericValue)) {
                            candidates.Add((row, column, numericValue));
                        }
                    }
                }

                if (candidates.Count == 0) {
                    continue;
                }

                double min = candidates.Min(candidate => candidate.Value);
                double max = candidates.Max(candidate => candidate.Value);
                foreach (var candidate in candidates) {
                    (double startRatio, double ratio) = GetDataBarGeometry(candidate.Value, min, max);
                    dataBars[(candidate.Row, candidate.Column)] = new ConditionalDataBarCell(rule.DataBarColor!, startRatio, ratio);
                }
            }

            foreach (ExcelConditionalFormattingInfo rule in rules
                .Where(rule => string.Equals(rule.Type, "IconSet", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(rule.IconSet))
                .OrderByDescending(rule => rule.Priority)) {
                var candidates = new List<(int Row, int Column, double Value)>();
                for (int row = 0; row < values.GetLength(0); row++) {
                    for (int column = 0; column < values.GetLength(1); column++) {
                        string? cellReference = cellReferences[row, column];
                        if (!string.IsNullOrWhiteSpace(cellReference) &&
                            IsCellReferenceInReferenceList(cellReference!, rule.Range) &&
                            TryGetConditionalNumericValue(values[row, column], out double numericValue)) {
                            candidates.Add((row, column, numericValue));
                        }
                    }
                }

                if (candidates.Count == 0) {
                    continue;
                }

                int iconCount = GetExcelIconSetCount(rule.IconSet!);
                double min = candidates.Min(candidate => candidate.Value);
                double max = candidates.Max(candidate => candidate.Value);
                foreach (var candidate in candidates) {
                    int bucket = GetExcelIconSetBucket(candidate.Value, min, max, iconCount);
                    if (rule.IconSetReverse) {
                        bucket = iconCount - 1 - bucket;
                    }

                    icons[(candidate.Row, candidate.Column)] = MapExcelIconSetCell(rule.IconSet!, bucket, iconCount);
                }
            }

            return fills.Count == 0 && dataBars.Count == 0 && icons.Count == 0 ? null : new ConditionalFillData(fills, dataBars, icons);
        }

        private static int GetExcelIconSetCount(string iconSet) {
            if (iconSet.StartsWith("Three", StringComparison.OrdinalIgnoreCase) ||
                iconSet.StartsWith("3", StringComparison.Ordinal)) {
                return 3;
            }

            if (iconSet.StartsWith("Four", StringComparison.OrdinalIgnoreCase) ||
                iconSet.StartsWith("4", StringComparison.Ordinal)) {
                return 4;
            }

            return 5;
        }

        private static int GetExcelIconSetBucket(double value, double min, double max, int iconCount) {
            if (iconCount <= 1 || max <= min) {
                return iconCount - 1;
            }

            double ratio = Math.Max(0D, Math.Min(1D, (value - min) / (max - min)));
            return Math.Max(0, Math.Min(iconCount - 1, (int)Math.Floor(ratio * iconCount)));
        }

        private static ConditionalIconCell MapExcelIconSetCell(string iconSet, int bucket, int iconCount) {
            string normalized = iconSet.ToLowerInvariant();
            bool trafficLights = normalized.IndexOf("traffic", StringComparison.Ordinal) >= 0;
            bool arrows = normalized.IndexOf("arrow", StringComparison.Ordinal) >= 0;
            bool symbols = normalized.IndexOf("symbol", StringComparison.Ordinal) >= 0 || normalized.IndexOf("sign", StringComparison.Ordinal) >= 0 || normalized.IndexOf("indicator", StringComparison.Ordinal) >= 0;

            if (trafficLights) {
                return new ConditionalIconCell(PdfCore.PdfCellIconKind.Circle, GetExcelIconBucketColor(bucket, iconCount));
            }

            if (arrows) {
                if (bucket == 0) {
                    return new ConditionalIconCell(PdfCore.PdfCellIconKind.TriangleDown, PdfCore.PdfColor.FromRgb(192, 80, 77));
                }

                if (bucket >= iconCount - 1) {
                    return new ConditionalIconCell(PdfCore.PdfCellIconKind.TriangleUp, PdfCore.PdfColor.FromRgb(99, 155, 71));
                }

                return new ConditionalIconCell(PdfCore.PdfCellIconKind.TriangleRight, PdfCore.PdfColor.FromRgb(255, 192, 0));
            }

            if (symbols && bucket == 0) {
                return new ConditionalIconCell(PdfCore.PdfCellIconKind.Diamond, PdfCore.PdfColor.FromRgb(192, 80, 77));
            }

            if (symbols && bucket >= iconCount - 1) {
                return new ConditionalIconCell(PdfCore.PdfCellIconKind.Circle, PdfCore.PdfColor.FromRgb(99, 155, 71));
            }

            if (symbols) {
                return new ConditionalIconCell(PdfCore.PdfCellIconKind.TriangleUp, PdfCore.PdfColor.FromRgb(255, 192, 0));
            }

            return new ConditionalIconCell(PdfCore.PdfCellIconKind.Circle, GetExcelIconBucketColor(bucket, iconCount));
        }

        private static PdfCore.PdfColor GetExcelIconBucketColor(int bucket, int iconCount) {
            if (bucket <= 0) {
                return PdfCore.PdfColor.FromRgb(192, 80, 77);
            }

            if (bucket >= iconCount - 1) {
                return PdfCore.PdfColor.FromRgb(99, 155, 71);
            }

            return PdfCore.PdfColor.FromRgb(255, 192, 0);
        }

        private static bool IsCellReferenceInReferenceList(string cellReference, string referenceList) {
            if (string.IsNullOrWhiteSpace(referenceList)) {
                return false;
            }

            (int Row, int Col) cell = A1.ParseCellRef(NormalizeCellReference(cellReference));
            if (cell.Row <= 0 || cell.Col <= 0) {
                return false;
            }

            foreach (string rawToken in referenceList.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                string token = StripSheetPrefix(rawToken).Replace("$", string.Empty);
                if (A1.TryParseRange(token, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                    if (cell.Row >= firstRow && cell.Row <= lastRow && cell.Col >= firstColumn && cell.Col <= lastColumn) {
                        return true;
                    }
                } else {
                    (int Row, int Col) singleCell = A1.ParseCellRef(token);
                    if (singleCell.Row == cell.Row && singleCell.Col == cell.Col) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool TryGetConditionalNumericValue(object? value, out double numericValue) {
            if (value is DateTime dateTime) {
                numericValue = dateTime.ToOADate();
                return true;
            }

            if (value is IConvertible convertible) {
                try {
                    numericValue = convertible.ToDouble(CultureInfo.InvariantCulture);
                    return !double.IsNaN(numericValue) && !double.IsInfinity(numericValue);
                } catch (FormatException) {
                } catch (InvalidCastException) {
                } catch (OverflowException) {
                }
            }

            numericValue = 0D;
            return false;
        }

        private static bool TryGetRgb(string value, out byte r, out byte g, out byte b) {
            string normalized = value.Trim().TrimStart('#');
            if (normalized.Length == 8) {
                normalized = normalized.Substring(2);
            }

            if (normalized.Length != 6 ||
                !byte.TryParse(normalized.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out r) ||
                !byte.TryParse(normalized.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out g) ||
                !byte.TryParse(normalized.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out b)) {
                r = 0;
                g = 0;
                b = 0;
                return false;
            }

            return true;
        }

        private static string InterpolateRgbHex(byte startR, byte startG, byte startB, byte endR, byte endG, byte endB, double ratio) {
            byte r = InterpolateByte(startR, endR, ratio);
            byte g = InterpolateByte(startG, endG, ratio);
            byte b = InterpolateByte(startB, endB, ratio);
            return r.ToString("X2", CultureInfo.InvariantCulture) +
                g.ToString("X2", CultureInfo.InvariantCulture) +
                b.ToString("X2", CultureInfo.InvariantCulture);
        }

        private static byte InterpolateByte(byte start, byte end, double ratio) {
            return (byte)Math.Max(0, Math.Min(255, (int)Math.Round(start + ((end - start) * ratio), MidpointRounding.AwayFromZero)));
        }

        private static VisibilityLayoutData? ReadVisibilityLayoutData(ExcelSheet? workbookSheet, string normalizedRange, int rowCount, int columnCount, bool enabled) {
            if (!enabled || workbookSheet == null || rowCount == 0 || columnCount == 0) {
                return null;
            }

            if (!A1.TryParseRange(normalizedRange, out int firstRow, out int firstColumn, out _, out _)) {
                return null;
            }

            IReadOnlyList<ExcelRowSnapshot> rowDefinitions = workbookSheet.GetRowDefinitions();
            IReadOnlyList<ExcelColumnSnapshot> columnDefinitions = workbookSheet.GetColumnDefinitions();
            if (!rowDefinitions.Any(row => row.Hidden) && !columnDefinitions.Any(column => column.Hidden)) {
                return null;
            }

            var rowOffsets = new List<int>(rowCount);
            for (int row = 0; row < rowCount; row++) {
                if (!IsWorksheetRowHidden(rowDefinitions, firstRow + row)) {
                    rowOffsets.Add(row);
                }
            }

            var columnOffsets = new List<int>(columnCount);
            for (int column = 0; column < columnCount; column++) {
                if (!IsWorksheetColumnHidden(columnDefinitions, firstColumn + column)) {
                    columnOffsets.Add(column);
                }
            }

            if (rowOffsets.Count == rowCount && columnOffsets.Count == columnCount) {
                return null;
            }

            return new VisibilityLayoutData(rowOffsets, columnOffsets, rowCount, columnCount);
        }

        private static string?[,]? ReadCellReferenceData(string normalizedRange, int rowCount, int columnCount, VisibilityLayoutData? visibility = null) {
            if (rowCount == 0 || columnCount == 0 ||
                !A1.TryParseRange(normalizedRange, out int firstRow, out int firstColumn, out _, out _)) {
                return null;
            }

            var references = new string?[rowCount, columnCount];
            for (int row = 0; row < rowCount; row++) {
                int sourceRow = visibility?.RowOffsets[row] ?? row;
                for (int column = 0; column < columnCount; column++) {
                    int sourceColumn = visibility?.ColumnOffsets[column] ?? column;
                    references[row, column] = A1.CellReference(firstRow + sourceRow, firstColumn + sourceColumn);
                }
            }

            return references;
        }

        private static object?[,] FilterValues(object?[,] values, VisibilityLayoutData? visibility) {
            if (visibility == null) {
                return values;
            }

            var result = new object?[visibility.RowOffsets.Count, visibility.ColumnOffsets.Count];
            for (int row = 0; row < visibility.RowOffsets.Count; row++) {
                for (int column = 0; column < visibility.ColumnOffsets.Count; column++) {
                    result[row, column] = values[visibility.RowOffsets[row], visibility.ColumnOffsets[column]];
                }
            }

            return result;
        }

        private static bool IsWorksheetRowHidden(IReadOnlyList<ExcelRowSnapshot> rowDefinitions, int rowIndex) {
            for (int i = rowDefinitions.Count - 1; i >= 0; i--) {
                ExcelRowSnapshot definition = rowDefinitions[i];
                if (definition.Index == rowIndex) {
                    return definition.Hidden;
                }
            }

            return false;
        }

        private static bool IsWorksheetColumnHidden(IReadOnlyList<ExcelColumnSnapshot> columnDefinitions, int columnIndex) {
            for (int i = columnDefinitions.Count - 1; i >= 0; i--) {
                ExcelColumnSnapshot definition = columnDefinitions[i];
                if (columnIndex >= definition.StartIndex && columnIndex <= definition.EndIndex) {
                    return definition.Hidden;
                }
            }

            return false;
        }

        private static ExcelCellStyleSnapshot?[,]? ReadCellStyleData(ExcelSheet? workbookSheet, string normalizedRange, int rowCount, int columnCount, bool enabled, VisibilityLayoutData? visibility = null) {
            if (!enabled || workbookSheet == null || rowCount == 0 || columnCount == 0) {
                return null;
            }

            if (!A1.TryParseRange(normalizedRange, out int firstRow, out int firstColumn, out _, out _)) {
                return null;
            }

            ExcelCellStyleSnapshot?[,] styles = new ExcelCellStyleSnapshot?[rowCount, columnCount];
            bool hasAnyStyle = false;
            for (int row = 0; row < rowCount; row++) {
                for (int column = 0; column < columnCount; column++) {
                    int sourceRow = visibility?.RowOffsets[row] ?? row;
                    int sourceColumn = visibility?.ColumnOffsets[column] ?? column;
                    ExcelCellStyleSnapshot style = workbookSheet.GetCellStyle(firstRow + sourceRow, firstColumn + sourceColumn);
                    if (style.HasPdfVisualStyle) {
                        styles[row, column] = style;
                        hasAnyStyle = true;
                    }
                }
            }

            return hasAnyStyle ? styles : null;
        }

        private static ExcelHyperlinkSnapshot?[,]? ReadHyperlinkData(ExcelSheet? workbookSheet, string normalizedRange, int rowCount, int columnCount, bool enabled, VisibilityLayoutData? visibility = null) {
            if (!enabled || workbookSheet == null || rowCount == 0 || columnCount == 0) {
                return null;
            }

            if (!A1.TryParseRange(normalizedRange, out int firstRow, out int firstColumn, out _, out _)) {
                return null;
            }

            IReadOnlyDictionary<string, ExcelHyperlinkSnapshot> worksheetHyperlinks = workbookSheet.GetHyperlinks();
            if (worksheetHyperlinks.Count == 0) {
                return null;
            }

            var links = new ExcelHyperlinkSnapshot?[rowCount, columnCount];
            bool hasAnyLink = false;
            for (int row = 0; row < rowCount; row++) {
                int sourceRow = visibility?.RowOffsets[row] ?? row;
                for (int column = 0; column < columnCount; column++) {
                    int sourceColumn = visibility?.ColumnOffsets[column] ?? column;
                    string reference = A1.CellReference(firstRow + sourceRow, firstColumn + sourceColumn);
                    if (TryGetHyperlink(worksheetHyperlinks, reference, out ExcelHyperlinkSnapshot? hyperlink) &&
                        IsSupportedPdfHyperlink(hyperlink, workbookSheet.Name)) {
                        links[row, column] = hyperlink;
                        hasAnyLink = true;
                    }
                }
            }

            return hasAnyLink ? links : null;
        }

        private static (double StartRatio, double Ratio) GetDataBarGeometry(double value, double min, double max) {
            if (max <= min) {
                return value < 0D ? (0D, 1D) : (0D, 1D);
            }

            if (min < 0D && max > 0D) {
                double range = max - min;
                double zeroRatio = Math.Max(0D, Math.Min(1D, -min / range));
                if (value >= 0D) {
                    return (zeroRatio, Math.Max(0D, Math.Min(1D - zeroRatio, value / range)));
                }

                double ratio = Math.Max(0D, Math.Min(zeroRatio, -value / range));
                return (zeroRatio - ratio, ratio);
            }

            if (max <= 0D) {
                double maxMagnitude = Math.Max(Math.Abs(min), Math.Abs(max));
                double ratio = maxMagnitude <= 0D ? 0D : Math.Max(0D, Math.Min(1D, Math.Abs(value) / maxMagnitude));
                return (1D - ratio, ratio);
            }

            double positiveRatio = Math.Max(0D, Math.Min(1D, (value - min) / (max - min)));
            return (0D, positiveRatio);
        }

        private static bool IsSupportedPdfHyperlink(ExcelHyperlinkSnapshot hyperlink, string currentSheetName) {
            if (hyperlink.IsExternal) {
                return Uri.TryCreate(hyperlink.Target, UriKind.Absolute, out _);
            }

            return TryParseInternalSheetName(hyperlink.Target, currentSheetName, out _);
        }

        private static bool TryGetHyperlink(IReadOnlyDictionary<string, ExcelHyperlinkSnapshot> hyperlinks, string cellReference, out ExcelHyperlinkSnapshot hyperlink) {
            if (hyperlinks.TryGetValue(cellReference, out ExcelHyperlinkSnapshot? direct)) {
                hyperlink = direct;
                return true;
            }

            foreach (KeyValuePair<string, ExcelHyperlinkSnapshot> entry in hyperlinks) {
                (int Row, int Col) cell = A1.ParseCellRef(cellReference);
                if (A1.TryParseRange(entry.Key, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn) &&
                    cell.Row >= firstRow &&
                    cell.Row <= lastRow &&
                    cell.Col >= firstColumn &&
                    cell.Col <= lastColumn) {
                    hyperlink = entry.Value;
                    return true;
                }
            }

            hyperlink = null!;
            return false;
        }

        private static ColumnLayoutData? ReadColumnLayoutData(ExcelSheet? workbookSheet, string normalizedRange, int columnCount, bool enabled, VisibilityLayoutData? visibility = null) {
            if (!enabled || workbookSheet == null || columnCount == 0) {
                return null;
            }

            if (!A1.TryParseRange(normalizedRange, out _, out int firstColumn, out _, out _)) {
                return null;
            }

            IReadOnlyList<ExcelColumnSnapshot> columnDefinitions = workbookSheet.GetColumnDefinitions();
            if (columnDefinitions.Count == 0) {
                return null;
            }

            var weights = new List<double>(columnCount);
            bool hasCustomWidth = false;
            double totalWidth = 0D;
            for (int columnOffset = 0; columnOffset < columnCount; columnOffset++) {
                int sourceColumnOffset = visibility?.ColumnOffsets[columnOffset] ?? columnOffset;
                int absoluteColumn = firstColumn + sourceColumnOffset;
                double width = GetWorksheetColumnWidth(columnDefinitions, absoluteColumn, out bool customWidth);
                weights.Add(width);
                totalWidth += width;
                hasCustomWidth |= customWidth;
            }

            return hasCustomWidth ? new ColumnLayoutData(weights, totalWidth * 5.25D) : null;
        }

        private static double GetWorksheetColumnWidth(IReadOnlyList<ExcelColumnSnapshot> columnDefinitions, int columnIndex, out bool customWidth) {
            for (int i = columnDefinitions.Count - 1; i >= 0; i--) {
                ExcelColumnSnapshot definition = columnDefinitions[i];
                if (columnIndex >= definition.StartIndex && columnIndex <= definition.EndIndex) {
                    customWidth = definition.CustomWidth && definition.Width.HasValue && definition.Width.Value > 0D;
                    return customWidth ? definition.Width!.Value : 8.43D;
                }
            }

            customWidth = false;
            return 8.43D;
        }

        private static RowLayoutData? ReadRowLayoutData(ExcelSheet? workbookSheet, string normalizedRange, int rowCount, bool enabled, VisibilityLayoutData? visibility = null) {
            if (!enabled || workbookSheet == null || rowCount == 0) {
                return null;
            }

            if (!A1.TryParseRange(normalizedRange, out int firstRow, out _, out _, out _)) {
                return null;
            }

            IReadOnlyList<ExcelRowSnapshot> rowDefinitions = workbookSheet.GetRowDefinitions();
            if (rowDefinitions.Count == 0) {
                return null;
            }

            var minHeights = new List<double?>(rowCount);
            bool hasCustomHeight = false;
            for (int rowOffset = 0; rowOffset < rowCount; rowOffset++) {
                int sourceRowOffset = visibility?.RowOffsets[rowOffset] ?? rowOffset;
                int absoluteRow = firstRow + sourceRowOffset;
                double? height = GetWorksheetRowHeight(rowDefinitions, absoluteRow);
                minHeights.Add(height);
                hasCustomHeight |= height.HasValue;
            }

            return hasCustomHeight ? new RowLayoutData(minHeights) : null;
        }

        private static double? GetWorksheetRowHeight(IReadOnlyList<ExcelRowSnapshot> rowDefinitions, int rowIndex) {
            for (int i = rowDefinitions.Count - 1; i >= 0; i--) {
                ExcelRowSnapshot definition = rowDefinitions[i];
                if (definition.Index == rowIndex) {
                    return definition.CustomHeight && definition.Height.HasValue && definition.Height.Value > 0D
                        ? definition.Height.Value
                        : null;
                }
            }

            return null;
        }

        private static MergeLayoutData? ReadMergeLayoutData(ExcelSheet? workbookSheet, string normalizedRange, int rowCount, int columnCount, bool enabled, VisibilityLayoutData? visibility = null) {
            if (!enabled || workbookSheet == null || rowCount == 0 || columnCount == 0) {
                return null;
            }

            if (!A1.TryParseRange(normalizedRange, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                return null;
            }

            var layout = new MergeLayoutData(rowCount, columnCount);
            foreach (ExcelMergedRangeSnapshot mergedRange in workbookSheet.GetMergedRanges()) {
                if (mergedRange.StartRow < firstRow ||
                    mergedRange.StartColumn < firstColumn ||
                    mergedRange.EndRow > lastRow ||
                    mergedRange.EndColumn > lastColumn) {
                    continue;
                }

                List<int> visibleRows = MapVisibleOffsets(mergedRange.StartRow - firstRow, mergedRange.EndRow - firstRow, visibility?.RowOffsets);
                List<int> visibleColumns = MapVisibleOffsets(mergedRange.StartColumn - firstColumn, mergedRange.EndColumn - firstColumn, visibility?.ColumnOffsets);
                if (visibleRows.Count == 0 || visibleColumns.Count == 0) {
                    continue;
                }

                int relativeRow = visibleRows[0];
                int relativeColumn = visibleColumns[0];
                int rowSpan = visibleRows.Count;
                int columnSpan = visibleColumns.Count;
                if (rowSpan > 1 || columnSpan > 1) {
                    layout.SetSpan(relativeRow, relativeColumn, rowSpan, columnSpan);
                }
            }

            return layout.HasAny ? layout : null;
        }

        private static List<int> MapVisibleOffsets(int firstSourceOffset, int lastSourceOffset, IReadOnlyList<int>? visibleOffsets) {
            if (visibleOffsets == null) {
                var all = new List<int>(lastSourceOffset - firstSourceOffset + 1);
                for (int offset = firstSourceOffset; offset <= lastSourceOffset; offset++) {
                    all.Add(offset);
                }

                return all;
            }

            var mapped = new List<int>();
            for (int index = 0; index < visibleOffsets.Count; index++) {
                int sourceOffset = visibleOffsets[index];
                if (sourceOffset >= firstSourceOffset && sourceOffset <= lastSourceOffset) {
                    mapped.Add(index);
                }
            }

            return mapped;
        }

        private static object?[,] PrependRows(object?[,] topRows, object?[,] bodyRows) {
            int topRowCount = topRows.GetLength(0);
            int bodyRowCount = bodyRows.GetLength(0);
            int columnCount = bodyRows.GetLength(1);
            var result = new object?[topRowCount + bodyRowCount, columnCount];
            CopyRows(topRows, result, 0, columnCount);
            CopyRows(bodyRows, result, topRowCount, columnCount);
            return result;
        }

        private static ExcelCellStyleSnapshot?[,]? PrependRows(ExcelCellStyleSnapshot?[,]? topRows, ExcelCellStyleSnapshot?[,]? bodyRows, int topRowCount, int bodyRowCount, int columnCount) {
            if (topRows == null && bodyRows == null) {
                return null;
            }

            var result = new ExcelCellStyleSnapshot?[topRowCount + bodyRowCount, columnCount];
            if (topRows != null) {
                CopyRows(topRows, result, 0, columnCount);
            }

            if (bodyRows != null) {
                CopyRows(bodyRows, result, topRowCount, columnCount);
            }

            return result;
        }

        private static ExcelHyperlinkSnapshot?[,]? PrependRows(ExcelHyperlinkSnapshot?[,]? topRows, ExcelHyperlinkSnapshot?[,]? bodyRows, int topRowCount, int bodyRowCount, int columnCount) {
            if (topRows == null && bodyRows == null) {
                return null;
            }

            var result = new ExcelHyperlinkSnapshot?[topRowCount + bodyRowCount, columnCount];
            if (topRows != null) {
                CopyRows(topRows, result, 0, columnCount);
            }

            if (bodyRows != null) {
                CopyRows(bodyRows, result, topRowCount, columnCount);
            }

            return result;
        }

        private static string?[,]? PrependRows(string?[,]? topRows, string?[,]? bodyRows, int topRowCount, int bodyRowCount, int columnCount) {
            if (topRows == null && bodyRows == null) {
                return null;
            }

            var result = new string?[topRowCount + bodyRowCount, columnCount];
            if (topRows != null) {
                CopyRows(topRows, result, 0, columnCount);
            }

            if (bodyRows != null) {
                CopyRows(bodyRows, result, topRowCount, columnCount);
            }

            return result;
        }

        private static MergeLayoutData? PrependRows(MergeLayoutData? topRows, MergeLayoutData? bodyRows, int topRowCount, int bodyRowCount, int columnCount) {
            if (topRows == null && bodyRows == null) {
                return null;
            }

            var result = new MergeLayoutData(topRowCount + bodyRowCount, columnCount);
            topRows?.CopyTo(result, 0);
            bodyRows?.CopyTo(result, topRowCount);
            return result.HasAny ? result : null;
        }

        private static RowLayoutData? PrependRows(RowLayoutData? topRows, RowLayoutData? bodyRows, int topRowCount, int bodyRowCount) {
            if (topRows == null && bodyRows == null) {
                return null;
            }

            var minHeights = new List<double?>(topRowCount + bodyRowCount);
            if (topRows != null) {
                minHeights.AddRange(topRows.MinHeights);
            } else {
                for (int row = 0; row < topRowCount; row++) {
                    minHeights.Add(null);
                }
            }

            if (bodyRows != null) {
                minHeights.AddRange(bodyRows.MinHeights);
            } else {
                for (int row = 0; row < bodyRowCount; row++) {
                    minHeights.Add(null);
                }
            }

            return minHeights.Any(height => height.HasValue) ? new RowLayoutData(minHeights) : null;
        }

        private static void CopyRows(object?[,] source, object?[,] target, int targetRowOffset, int columnCount) {
            int rowCount = source.GetLength(0);
            int sourceColumnCount = source.GetLength(1);
            for (int row = 0; row < rowCount; row++) {
                for (int column = 0; column < columnCount; column++) {
                    target[targetRowOffset + row, column] = column < sourceColumnCount ? source[row, column] : null;
                }
            }
        }

        private static void CopyRows(ExcelCellStyleSnapshot?[,] source, ExcelCellStyleSnapshot?[,] target, int targetRowOffset, int columnCount) {
            int rowCount = source.GetLength(0);
            int sourceColumnCount = source.GetLength(1);
            for (int row = 0; row < rowCount; row++) {
                for (int column = 0; column < columnCount; column++) {
                    target[targetRowOffset + row, column] = column < sourceColumnCount ? source[row, column] : null;
                }
            }
        }

        private static void CopyRows(ExcelHyperlinkSnapshot?[,] source, ExcelHyperlinkSnapshot?[,] target, int targetRowOffset, int columnCount) {
            int rowCount = source.GetLength(0);
            int sourceColumnCount = source.GetLength(1);
            for (int row = 0; row < rowCount; row++) {
                for (int column = 0; column < columnCount; column++) {
                    target[targetRowOffset + row, column] = column < sourceColumnCount ? source[row, column] : null;
                }
            }
        }

        private static void CopyRows(string?[,] source, string?[,] target, int targetRowOffset, int columnCount) {
            int rowCount = source.GetLength(0);
            int sourceColumnCount = source.GetLength(1);
            for (int row = 0; row < rowCount; row++) {
                for (int column = 0; column < columnCount; column++) {
                    target[targetRowOffset + row, column] = column < sourceColumnCount ? source[row, column] : null;
                }
            }
        }

        private static string NormalizeA1Range(string range) {
            string withoutSheet = StripSheetPrefix(range).Replace("$", string.Empty);
            if (!A1.TryParseRange(withoutSheet, out int r1, out int c1, out int r2, out int c2)) {
                (int Row, int Col) cell = A1.ParseCellRef(withoutSheet);
                if (cell.Row <= 0 || cell.Col <= 0) {
                    throw new ArgumentException("Excel PDF export range must be a valid A1 range.", nameof(range));
                }

                r1 = r2 = cell.Row;
                c1 = c2 = cell.Col;
            }

            return ToA1Range(r1, c1, r2, c2);
        }

        private static string StripSheetPrefix(string range) {
            int separator = range.LastIndexOf('!');
            return separator >= 0 ? range.Substring(separator + 1) : range;
        }

        private static string ToA1Range(int firstRow, int firstColumn, int lastRow, int lastColumn) {
            string start = A1.ColumnIndexToLetters(firstColumn) + firstRow.ToString(CultureInfo.InvariantCulture);
            string end = A1.ColumnIndexToLetters(lastColumn) + lastRow.ToString(CultureInfo.InvariantCulture);
            return start + ":" + end;
        }

    }
}

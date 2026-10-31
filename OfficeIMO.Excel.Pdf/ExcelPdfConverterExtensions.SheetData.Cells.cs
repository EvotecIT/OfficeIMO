using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static VisibilityLayoutData? ReadVisibilityLayoutData(ExcelSheet? workbookSheet, string normalizedRange, int rowCount, int columnCount, bool enabled, System.Threading.CancellationToken cancellationToken) {
            if (!enabled || workbookSheet == null || rowCount == 0 || columnCount == 0) {
                return null;
            }

            if (!ExcelA1.TryParseRange(normalizedRange, out int firstRow, out int firstColumn, out _, out _)) {
                return null;
            }

            IReadOnlyList<ExcelRowInfo> rowDefinitions = workbookSheet.GetRowDefinitions();
            IReadOnlyList<ExcelColumnInfo> columnDefinitions = workbookSheet.GetColumnDefinitions(cancellationToken);
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
                !ExcelA1.TryParseRange(normalizedRange, out int firstRow, out int firstColumn, out _, out _)) {
                return null;
            }

            var references = new string?[rowCount, columnCount];
            for (int row = 0; row < rowCount; row++) {
                int sourceRow = visibility?.RowOffsets[row] ?? row;
                for (int column = 0; column < columnCount; column++) {
                    int sourceColumn = visibility?.ColumnOffsets[column] ?? column;
                    references[row, column] = ExcelA1.CellReference(firstRow + sourceRow, firstColumn + sourceColumn);
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

        private static bool IsWorksheetRowHidden(IReadOnlyList<ExcelRowInfo> rowDefinitions, int rowIndex) {
            for (int i = rowDefinitions.Count - 1; i >= 0; i--) {
                ExcelRowInfo definition = rowDefinitions[i];
                if (definition.Index == rowIndex) {
                    return definition.Hidden;
                }
            }

            return false;
        }

        private static bool IsWorksheetColumnHidden(IReadOnlyList<ExcelColumnInfo> columnDefinitions, int columnIndex) {
            for (int i = columnDefinitions.Count - 1; i >= 0; i--) {
                ExcelColumnInfo definition = columnDefinitions[i];
                if (columnIndex >= definition.StartIndex && columnIndex <= definition.EndIndex) {
                    return definition.Hidden;
                }
            }

            return false;
        }

        private static ExcelCellStyleInfo?[,]? ReadCellStyleData(ExcelSheet? workbookSheet, ExcelSheetReader sheetReader, string normalizedRange, int rowCount, int columnCount, bool enabled, PdfCore.PdfStandardFont defaultFontFamily, bool boundedRead, VisibilityLayoutData? visibility = null) {
            if (!enabled || workbookSheet == null || rowCount == 0 || columnCount == 0) {
                return null;
            }

            if (!ExcelA1.TryParseRange(normalizedRange, out int firstRow, out int firstColumn, out _, out _)) {
                return null;
            }

            ExcelCellStyleInfo?[,] styles = new ExcelCellStyleInfo?[rowCount, columnCount];
            bool hasAnyStyle = false;
            uint?[,]? styleIndexes = boundedRead
                ? sheetReader.ReadCellStyleIndexes(normalizedRange)
                : null;
            bool useStreamedStyleIndexes = boundedRead && styleIndexes != null;
            for (int row = 0; row < rowCount; row++) {
                for (int column = 0; column < columnCount; column++) {
                    int sourceRow = visibility?.RowOffsets[row] ?? row;
                    int sourceColumn = visibility?.ColumnOffsets[column] ?? column;
                    uint? styleIndex = useStreamedStyleIndexes &&
                        styleIndexes != null &&
                        sourceRow < styleIndexes.GetLength(0) &&
                        sourceColumn < styleIndexes.GetLength(1)
                            ? styleIndexes[sourceRow, sourceColumn]
                            : null;
                    if (useStreamedStyleIndexes && !styleIndex.HasValue) {
                        continue;
                    }

                    ExcelCellStyleInfo style = useStreamedStyleIndexes
                        ? workbookSheet.GetCellStyleByIndex(styleIndex!.Value)
                        : workbookSheet.GetCellStyle(firstRow + sourceRow, firstColumn + sourceColumn);
                    if (HasPdfExportStyle(style, defaultFontFamily)) {
                        styles[row, column] = style;
                        hasAnyStyle = true;
                    }
                }
            }

            return hasAnyStyle ? styles : null;
        }

        private static bool HasPdfExportStyle(ExcelCellStyleInfo style, PdfCore.PdfStandardFont defaultFontFamily) {
            if (style.HasPdfVisualStyle) {
                return true;
            }

            return PdfCore.PdfStandardFontMapper.TryMapFontFamily(style.FontName, out PdfCore.PdfStandardFont font) &&
                PdfCore.PdfStandardFontMapper.GetFontFamily(font) != defaultFontFamily;
        }

        private static ExcelHyperlinkInfo?[,]? ReadHyperlinkData(ExcelSheet? workbookSheet, string normalizedRange, int rowCount, int columnCount, bool enabled, VisibilityLayoutData? visibility = null) {
            if (!enabled || workbookSheet == null || rowCount == 0 || columnCount == 0) {
                return null;
            }

            if (!ExcelA1.TryParseRange(normalizedRange, out int firstRow, out int firstColumn, out _, out _)) {
                return null;
            }

            IReadOnlyDictionary<string, ExcelHyperlinkInfo> worksheetHyperlinks = workbookSheet.GetHyperlinks();
            if (worksheetHyperlinks.Count == 0) {
                return null;
            }

            var links = new ExcelHyperlinkInfo?[rowCount, columnCount];
            bool hasAnyLink = false;
            for (int row = 0; row < rowCount; row++) {
                int sourceRow = visibility?.RowOffsets[row] ?? row;
                for (int column = 0; column < columnCount; column++) {
                    int sourceColumn = visibility?.ColumnOffsets[column] ?? column;
                    string reference = ExcelA1.CellReference(firstRow + sourceRow, firstColumn + sourceColumn);
                    if (TryGetHyperlink(worksheetHyperlinks, reference, out ExcelHyperlinkInfo? hyperlink) &&
                        IsSupportedPdfHyperlink(hyperlink, workbookSheet.Name)) {
                        links[row, column] = hyperlink;
                        hasAnyLink = true;
                    }
                }
            }

            return hasAnyLink ? links : null;
        }


        private static bool IsSupportedPdfHyperlink(ExcelHyperlinkInfo hyperlink, string currentSheetName) {
            if (hyperlink.IsExternal) {
                return Uri.TryCreate(hyperlink.Target, UriKind.Absolute, out _);
            }

            return TryParseInternalSheetName(hyperlink.Target, currentSheetName, out _);
        }

        private static bool TryGetHyperlink(IReadOnlyDictionary<string, ExcelHyperlinkInfo> hyperlinks, string cellReference, out ExcelHyperlinkInfo hyperlink) {
            if (hyperlinks.TryGetValue(cellReference, out ExcelHyperlinkInfo? direct)) {
                hyperlink = direct;
                return true;
            }

            foreach (KeyValuePair<string, ExcelHyperlinkInfo> entry in hyperlinks) {
                (int Row, int Col) cell = ExcelA1.ParseCellRef(cellReference);
                if (ExcelA1.TryParseRange(entry.Key, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn) &&
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

    }
}

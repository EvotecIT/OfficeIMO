using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
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

        private static ExcelCellStyleSnapshot?[,]? ReadCellStyleData(ExcelSheet? workbookSheet, string normalizedRange, int rowCount, int columnCount, bool enabled, PdfCore.PdfStandardFont defaultFontFamily, VisibilityLayoutData? visibility = null) {
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
                    if (HasPdfExportStyle(style, defaultFontFamily)) {
                        styles[row, column] = style;
                        hasAnyStyle = true;
                    }
                }
            }

            return hasAnyStyle ? styles : null;
        }

        private static bool HasPdfExportStyle(ExcelCellStyleSnapshot style, PdfCore.PdfStandardFont defaultFontFamily) {
            if (style.HasPdfVisualStyle) {
                return true;
            }

            return PdfCore.PdfStandardFontMapper.TryMapFontFamily(style.FontName, out PdfCore.PdfStandardFont font) &&
                PdfCore.PdfStandardFontMapper.GetFontFamily(font) != defaultFontFamily;
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

    }
}

using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
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
    }
}

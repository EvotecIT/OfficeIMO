namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Measures the fragment used by table fitting and painting, including per-cell content and padding.</summary>
        private static double MeasurePreparedTableRowSegmentHeight(TableBlock table, PdfTableStyle style,
            PreparedFlowTableRows rows, int columns, double[] columnWidths, double columnGap,
            int rowIndex, int startLine, int lineCount, bool suppressCellObjects) {
            if (startLine == 0 && lineCount == rows.LineCounts[rowIndex]) return rows.Heights[rowIndex];
            double leading = rows.Leadings[rowIndex];
            double height = leading + GetTableRowMaxPaddingTop(table, style, rowIndex, columns) +
                GetTableRowMaxPaddingBottom(table, style, rowIndex, columns);
            foreach (TableCellLayout cell in GetTableCellLayouts(table, rowIndex, columns)) {
                double cellWidth = GetTableCellWidth(columnWidths, cell.Column, cell.ColumnSpan, columnGap);
                double innerWidth = cellWidth - GetTableCellPaddingLeft(style, rowIndex, cell.Column) -
                    GetTableCellPaddingRight(style, rowIndex, cell.Column);
                TableCellTextLayout lines = rows.Lines[rowIndex][cell.Column];
                int visibleLines = Math.Max(0, Math.Min(lineCount, lines.LineCount - startLine));
                bool includeObjects = !suppressCellObjects && startLine == 0;
                double cellHeight = MeasureTableCellContentHeight(cell, lines, startLine, visibleLines,
                    leading, innerWidth, includeObjects) + GetTableCellPaddingTop(style, rowIndex, cell.Column) +
                    GetTableCellPaddingBottom(style, rowIndex, cell.Column);
                height = Math.Max(height, cellHeight);
            }
            return startLine == 0 ? Math.Max(height,
                GetTableRowFixedHeight(style, rowIndex) ?? GetTableRowMinHeight(style, rowIndex)) : height;
        }

        /// <summary>Fits a measured row fragment without crossing the cells' legal paragraph boundaries.</summary>
        private static int GetPreparedTableRowSegmentLineCountThatFits(TableBlock table, PdfTableStyle style,
            PreparedFlowTableRows rows, int columns, double[] columnWidths, double columnGap,
            int rowIndex, int startLine, double available, double fullFrameHeight, bool canMoveToNextFrame,
            bool requireDefaultFirstFragment = false, int? maximumLineCount = null) {
            int remaining = rows.LineCounts[rowIndex] - startLine;
            if (maximumLineCount.HasValue) remaining = Math.Min(remaining, maximumLineCount.Value);
            int best = 0;
            for (int candidate = 1; candidate <= remaining; candidate++) {
                double height = MeasurePreparedTableRowSegmentHeight(table, style, rows, columns,
                    columnWidths, columnGap, rowIndex, startLine, candidate, suppressCellObjects: false);
                if (height > available + 0.001D) break;
                best = candidate;
            }
            return LimitTableRowFragmentToParagraphBoundaries(rows.Lines[rowIndex],
                GetTableCellLayouts(table, rowIndex, columns), startLine, best,
                fullFrameHeight, canMoveToNextFrame, requireDefaultFirstFragment);
        }
    }
}

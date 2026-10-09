namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Preflights a breakable first row at its first legal cell boundary instead of requiring the entire row.</summary>
        private double MeasureTableFirstFragmentHeight(TableBlock table, PdfTableStyle style, PreparedFlowTableRows rows,
            int columns, double[] widths, double gap) {
            if (style.KeepTogether || !GetTableRowAllowBreakAcrossPages(style, 0) || TableRowHasViewport(table, 0, columns))
                return rows.Heights[0];
            IReadOnlyList<TableCellLayout> cells = GetTableCellLayouts(table, 0, columns);
            var spanFlow = new TableSpanFlow(table, style, columns, style.HeaderRowCount,
                table.Rows.Count - style.FooterRowCount);
            bool hasIndependentSpan = spanFlow.CoversRow(0);
            if (hasIndependentSpan)
                cells = cells.Where(cell => !spanFlow.Contains(0, cell.Column)).ToArray();
            double fullHeight = GetMeasuredContinuationFrameHeight();
            for (int count = 1; count <= rows.LineCounts[0]; count++) {
                if (IsTableRowFragmentBoundaryAllowed(rows.Lines[0], cells, 0, count, fullHeight, requireDefaultFirstFragment: true)) {
                    double height = MeasurePreparedTableRowSegmentHeight(table, style, rows, columns, widths, gap,
                        0, 0, count, suppressCellObjects: false);
                    return hasIndependentSpan ? Math.Min(height, rows.IntrinsicHeights[0]) : height;
                }
            }
            return rows.Heights[0];
        }
    }
}

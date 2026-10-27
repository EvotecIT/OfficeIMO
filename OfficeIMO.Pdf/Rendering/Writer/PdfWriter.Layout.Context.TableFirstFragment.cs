namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Previews only references in the first legal visual fragment used by a preceding keep chain.</summary>
        private IEnumerable<RichSeg> GetFirstTableFootnoteSegments(TableBlock table) {
            PdfTableStyle style = table.Style ?? currentOpts.DefaultTableStyleSnapshot ?? PdfTableStyles.Light();
            int columns = GetTableColumnCount(table);
            if (columns == 0 || table.Rows.Count == 0 || !ValidateTableFootnotePlacement(table, style, columns))
                return Enumerable.Empty<RichSeg>();
            style = PreparePairedTableBorders(table, style);
            double gap = GetTableCellSpacing(style);
            int footerStart = table.Rows.Count - style.FooterRowCount;
            double fontSize = GetTableBodyFontSize(style, currentOpts.DefaultFontSize);
            TableColumnLayout columnLayout = ResolveTableColumnLayout(table, currentOpts, style, columns,
                currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight,
                fontSize, style.HeaderRowCount, footerStart);
            PreparedFlowTableRows rows = PrepareFlowTableRows(table, style, columns, columnLayout.Widths,
                gap, gap, style.HeaderRowCount, footerStart);
            var spans = new TableSpanFlow(table, style, columns, style.HeaderRowCount, footerStart);
            double firstHeight = MeasureTableFirstFragmentHeight(table, style, rows, columns, columnLayout.Widths, gap);
            int count = rows.LineCounts[0];
            if (!style.KeepTogether && GetTableRowAllowBreakAcrossPages(style, 0) && !TableRowHasViewport(table, 0, columns)) {
                IReadOnlyList<TableCellLayout> cells = GetTableCellLayouts(table, 0, columns);
                if (spans.CoversRow(0)) cells = cells.Where(cell => !spans.Contains(0, cell.Column)).ToArray();
                for (int candidate = 1; candidate <= rows.LineCounts[0]; candidate++) {
                    if (!IsTableRowFragmentBoundaryAllowed(rows.Lines[0], cells, 0, candidate,
                        GetMeasuredContinuationFrameHeight(), requireDefaultFirstFragment: true)) continue;
                    count = candidate;
                    break;
                }
            }
            return GetTableRowFootnoteSegments(table, style, rows, columns, columnLayout.Widths, gap, gap,
                spans, 0, 0, count, y, firstHeight, currentOpts.MarginLeft,
                GetMeasuredContinuationFrameHeight(), GetClosingTextPadding(style.SpacingAfter));
        }

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

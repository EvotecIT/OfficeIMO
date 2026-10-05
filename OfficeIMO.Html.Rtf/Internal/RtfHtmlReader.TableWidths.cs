namespace OfficeIMO.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private int GetPageTextWidthTwips() {
            RtfPageSetup page = _document.PageSetup;
            RtfPageSetup? section = _currentSection?.PageSetup;
            long width = (long)(section?.PaperWidthTwips ?? page.PaperWidthTwips ?? 12240)
                - (section?.MarginLeftTwips ?? page.MarginLeftTwips ?? 1800)
                - (section?.MarginRightTwips ?? page.MarginRightTwips ?? 1800)
                - (section?.GutterWidthTwips ?? page.GutterWidthTwips ?? 0);
            return (int)Math.Max(1L, Math.Min(int.MaxValue, width));
        }

        private void FitDefaultTableColumns(RtfTable table, int availableWidth) {
            // Sparse rows and trailing rowspan continuations keep logical column
            // boundaries even when row.Cells has fewer entries than the grid.
            bool hasDefaultGrid = _automaticWidthTables.Contains(table) && table.Rows.Count > 0 &&
                table.Rows.All(row => !row.PreferredWidth.HasValue && row.Cells.All(cell =>
                    !cell.PreferredWidth.HasValue && cell.RightBoundaryTwips > 0 && cell.RightBoundaryTwips % 2400 == 0));
            int columns = hasDefaultGrid ? table.Rows.SelectMany(row => row.Cells)
                .Select(cell => cell.RightBoundaryTwips!.Value / 2400).DefaultIfEmpty().Max() : 0;
            if (columns > 0 && (long)columns * 2400 > availableWidth) {
                int columnWidth = Math.Max(1, availableWidth / columns);
                foreach (RtfTableRow row in table.Rows) {
                    foreach (RtfTableCell cell in row.Cells)
                        cell.RightBoundaryTwips = cell.RightBoundaryTwips!.Value / 2400 * columnWidth;
                }
            }
            // A containing table may be fitted after its children were parsed.
            // Revisit ordinary nested grids against the final host cell geometry.
            foreach (RtfTableRow row in table.Rows) {
                int left = 0;
                for (int i = 0; i < row.Cells.Count; i++) {
                    RtfTableCell cell = row.Cells[i];
                    int right = cell.RightBoundaryTwips ?? left;
                    int contentRight = right;
                    if (cell.HorizontalMerge == RtfTableCellMerge.First) {
                        for (int next = i + 1; next < row.Cells.Count &&
                            row.Cells[next].HorizontalMerge == RtfTableCellMerge.Continue; next++)
                            contentRight = row.Cells[next].RightBoundaryTwips ?? contentRight;
                    }
                    long contentWidth = (long)contentRight - left
                        - (cell.PaddingLeftTwips ?? row.PaddingLeftTwips ?? 108)
                        - (cell.PaddingRightTwips ?? row.PaddingRightTwips ?? 108);
                    int boundedWidth = (int)Math.Max(1L, Math.Min(int.MaxValue, contentWidth));
                    FitIntrinsicCellImages(cell, boundedWidth);
                    foreach (RtfTable nested in cell.Blocks.OfType<RtfTable>()) FitDefaultTableColumns(nested, boundedWidth);
                    left = right;
                }
            }
        }
    }
}

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private List<ColumnBalanceUnit>? MeasureTableColumnBalanceUnits(TableBlock table, ColumnFlowScope scope, double frameWidth) {
            PdfTableStyle style = table.Style ?? currentOpts.DefaultTableStyleSnapshot ?? TableStyles.Light();
            int columns = GetTableColumnCount(table);
            if (columns == 0 || table.Rows.Count == 0) return new();
            if (style.Position != null || !style.ConsumesVerticalFlow) return null;
            ValidateTableRoleRowCounts(style, table.Rows.Count);
            ValidateTableCellStyleCoordinates(style, table, columns);
            ValidateTableColumnStyleBounds(style, columns);
            ValidateTableRowStyleBounds(style, table.Rows.Count);
            int footerStart = table.Rows.Count - style.FooterRowCount;
            ValidateTableRowSpansWithinRoleBoundaries(table, columns, style.HeaderRowCount, footerStart);
            double gap = GetTableCellSpacing(style);
            double fontSize = GetTableBodyFontSize(style, currentOpts.DefaultFontSize);
            TableColumnLayout layout = ResolveTableColumnLayout(table, currentOpts, style, columns, frameWidth, fontSize,
                style.HeaderRowCount, footerStart);
            PreparedFlowTableRows rows = PrepareFlowTableRows(table, style, columns, layout.Widths, gap, gap, style.HeaderRowCount, footerStart);
            double before = ResolveTopLevelSpacingBefore(style.SpacingBefore);
            if (!string.IsNullOrWhiteSpace(style.Caption)) {
                double size = style.CaptionFontSize ?? fontSize;
                double leading = size * 1.25D;
                var wrapped = WrapRichRunsCore(new[] { PdfTextRun.Normal(style.Caption!, style.CaptionColor, size) }, layout.Width,
                    size, ChooseNormal(currentOpts.DefaultFont), leading, null, DefaultParagraphTabStopWidth, currentOpts);
                before += MeasureRichLinesHeight(wrapped.LineHeights, wrapped.Lines.Count, leading) + style.CaptionSpacingAfter;
            }
            return MeasurePreparedTableColumnBalanceUnits(table, style, rows, columns, layout.Widths, gap,
                GetMeasuredContinuationFrameHeight(), scope.Options.BalanceTableRowLines, initialHeight: before);
        }

        /// <summary>Measures row groups and optional cell fragments with the existing header, body and viewport constraints.</summary>
        private static List<ColumnBalanceUnit>? MeasurePreparedTableColumnBalanceUnits(TableBlock table, PdfTableStyle style,
            PreparedFlowTableRows rows, int columns, double[] columnWidths, double gap, double fullFrameHeight,
            bool balanceRowLines, int firstRow = 0, int consumedLines = 0, double initialHeight = 0D) {
            if ((!balanceRowLines && consumedLines > 0) || style.Position != null || !style.ConsumesVerticalFlow) return null;
            var units = new List<ColumnBalanceUnit>();
            int headerCount = style.HeaderRowCount;
            int footerStart = table.Rows.Count - style.FooterRowCount;
            int repeatHeaders = GetTableRepeatHeaderRowCount(style);
            double repeatHeight = table.Rows.Count > headerCount ? GetTableRowsHeight(rows.Heights, 0, repeatHeaders, gap) : 0D;
            int[] viewportGroups = GetTableViewportRowGroups(table, columns);
            for (int row = firstRow; row < table.Rows.Count;) {
                int end = row;
                if (row == 0) end = Math.Max(end, Math.Min(footerStart - 1, headerCount + Math.Max(1, style.MinimumBodyRowsOnFirstPage) - 1));
                int finalStart = Math.Max(headerCount, footerStart - style.MinimumBodyRowsOnLastPage);
                if (row >= finalStart) end = table.Rows.Count - 1;
                for (int member = row; member <= end; member++) end = Math.Max(end, viewportGroups[member]);
                double height = GetTableRowsHeight(rows.Heights, row, end - row + 1, gap);
                // The range helper includes a trailing gap whenever the range ends before the table.
                double continuation = (row > 0 || consumedLines > 0) && row >= headerCount
                    ? repeatHeight + style.PageContinuationSpacingBefore : 0D;
                double before = row == firstRow ? (firstRow == 0 && consumedLines == 0 ? initialHeight : continuation) : 0D;
                double after = end == table.Rows.Count - 1 ? style.SpacingAfter : 0D;
                int fragmentRow = -1;
                if (balanceRowLines && !style.KeepTogether) {
                    int fragmentCount = 0;
                    for (int member = row; member <= end; member++) {
                        if (rows.LineCounts[member] <= 1 ||
                            !GetTableRowAllowBreakAcrossPages(style, member) || TableRowHasViewport(table, member, columns) ||
                            viewportGroups[member] >= member) continue;
                        fragmentCount++;
                        if (fragmentRow < 0) fragmentRow = member;
                    }
                    if (fragmentCount > 1) {
                        // Ordinary table flow relaxes minimum body-row groups that exceed a full frame.
                        // Viewport groups remain indivisible even when their text rows could otherwise split.
                        bool hasViewport = Enumerable.Range(row, end - row + 1).Any(member => viewportGroups[member] >= member);
                        if (height > fullFrameHeight + 0.001D && !hasViewport) {
                            end = fragmentRow;
                            after = end == table.Rows.Count - 1 ? style.SpacingAfter : 0D;
                        } else fragmentRow = -1;
                    }
                }
                if (fragmentRow >= 0) {
                    double prefix = GetTableRowsHeight(rows.Heights, row, fragmentRow - row, gap);
                    double suffix = GetTableRowsHeight(rows.Heights, fragmentRow + 1, end - fragmentRow, gap) + after +
                        GetTableRowGapAfter(fragmentRow, table.Rows.Count, gap);
                    int startLine = fragmentRow == firstRow ? consumedLines : 0;
                    double splitBefore = (fragmentRow >= headerCount ? repeatHeight : 0D) + style.PageContinuationSpacingBefore;
                    if (row == firstRow && consumedLines > 0) before = splitBefore;
                    units.Add(new ColumnBalanceUnit(new ColumnBalanceRowFragment(table, style, rows, columns,
                        columnWidths, gap, fragmentRow, startLine, fullFrameHeight, prefix + before,
                        prefix + (row == firstRow ? before : continuation), splitBefore, suffix)));
                } else {
                    if (row == firstRow && consumedLines > 0) return null;
                    units.Add(new ColumnBalanceUnit(height + before + after, continuation));
                }
                row = end + 1;
            }
            if (style.KeepTogether && units.Count > 0) return new() { new(units.Sum(unit => unit.Height)) };
            return units;
        }
    }
}

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Prepares table cell text for the current frame, optionally retaining an unfinished row's cursor.</summary>
        private PreparedFlowTableRows PrepareFlowTableRows(TableBlock table, PdfTableStyle style, int columns,
            double[] columnWidths, double columnGap, double rowGap, int headerCount, int footerStart,
            PreparedFlowTableRows? previous = null, int continuingRow = -1, int consumedLines = 0, double? fallbackFontSize = null, TableSpanFlow? spanFlow = null) {
            var result = new PreparedFlowTableRows(table.Rows.Count);
            var textLayouts = new TableTextLayoutReuse(currentOpts);
            for (int row = 0; row < table.Rows.Count; row++) {
                bool continued = row == continuingRow && consumedLines > 0 && previous != null;
                double originalSize = GetTableRowFontSize(style, row, headerCount, footerStart, fallbackFontSize ?? currentOpts.DefaultFontSize);
                bool bold = GetTableRowBold(style, row, headerCount, footerStart);
                TableRowTextSizing sizing = ResolveTableRowTextSizing(table, style, row, columns, columnWidths, columnGap, originalSize, bold, currentOpts);
                double size = continued ? Math.Min(previous!.Sizes[row], sizing.FontSize) : sizing.FontSize;
                double runScale = continued ? Math.Min(previous!.RunFontSizeScales[row], sizing.RunFontSizeScale) : sizing.RunFontSizeScale;
                double continuationScale = continued ? Math.Min(size / previous!.Sizes[row], runScale / previous.RunFontSizeScales[row]) : 1D;
                double leading = continued && Math.Abs(size - previous!.Sizes[row]) <= .001D
                    ? previous.Leadings[row] : GetTableLeading(style, size);
                result.Sizes[row] = size;
                result.Leadings[row] = leading;
                result.RunFontSizeScales[row] = runScale;
                result.Bold[row] = bold;
                result.Lines[row] = new TableCellTextLayout[columns];
                int maxLines = 1;
                double maxHeight = GetTableRowInitialTextHeight(table, row, columns, leading) + GetTableRowMaxPaddingTop(table, style, row, columns) + GetTableRowMaxPaddingBottom(table, style, row, columns);
                for (int column = 0; column < columns; column++) {
                    result.Lines[row][column] = new TableCellTextLayout(new() { new() }, new() { leading });
                }
                foreach (TableCellLayout cell in GetTableCellLayouts(table, row, columns)) {
                    var font = GetTableRowFont(currentOpts, bold);
                    double cellWidth = GetTableCellWidth(columnWidths, cell.Column, cell.ColumnSpan, columnGap);
                    double innerWidth = Math.Max(1D, GetTableCellContentWidth(cell, cellWidth) -
                        GetTableCellPaddingLeft(style, row, cell.Column) - GetTableCellPaddingRight(style, row, cell.Column));
                    bool mergedCursor = spanFlow?.Contains(row, cell.Column) == true;
                    int spanConsumed = mergedCursor ? spanFlow!.GetConsumedLines(row, cell.Column) : 0;
                    bool continueCell = previous != null && (mergedCursor ? spanConsumed > 0 : continued);
                    int cellConsumed = mergedCursor ? spanConsumed : consumedLines;
                    TableCellTextLayout lines = continueCell
                        ? ContinueTableCellTextLayout(cell, previous!.Lines[row][cell.Column], cellConsumed, innerWidth, font, size, leading,
                            currentOpts, continuationScale, style.MinimumShrinkFontSize ?? 6D, style.AutoFitWidthUsesContentMinimum)
                        : textLayouts.Create(cell, innerWidth, font, size, leading, runScale, style.MinimumShrinkFontSize ?? 6D, style.AutoFitWidthUsesContentMinimum);
                    result.Lines[row][cell.Column] = lines;
                    if (cell.RowSpan <= 1 && cell.Viewport == null) {
                        maxLines = Math.Max(maxLines, cell.TextRotation == 0 ? lines.LineCount : 1);
                        maxHeight = Math.Max(maxHeight, MeasureTableCellContentHeight(cell, lines, consumedLines > 0 && continued ? consumedLines : 0,
                            continued ? Math.Max(0, lines.LineCount - consumedLines) : lines.LineCount, leading, innerWidth, includeObjects: !continued) +
                            GetTableCellPaddingTop(style, row, cell.Column) + GetTableCellPaddingBottom(style, row, cell.Column));
                    }
                }
                result.LineCounts[row] = maxLines;
                result.Heights[row] = ResolveTableRowHeight(style, row, maxHeight);
                result.IntrinsicHeights[row] = result.Heights[row];
            }
            ApplyTableRowSpanHeights(table, style, columns, columnWidths, result.Lines, result.Heights, result.Leadings, columnGap, rowGap);
            return result;
        }


    }
}

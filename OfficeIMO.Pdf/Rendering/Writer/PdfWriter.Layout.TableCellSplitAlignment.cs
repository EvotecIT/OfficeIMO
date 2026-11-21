namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>
        /// Lets each cell of a breaking row keep the lines that fit in the available height under its own
        /// paragraph rules, as Word does, instead of cutting every cell at the most restricted cell's line.
        /// </summary>
        /// <remarks>
        /// Row fragments address one line index shared by all cells. A cell that stops before the deepest
        /// cell receives zero-height placeholder lines up to that index, the same placeholders continuation
        /// layouts use for consumed lines, so the shared boundary then gives every cell its own break.
        /// Rows with merged, rotated or viewport cells, and rows that column balancing measures
        /// by shared lines, keep the shared boundary.
        /// </remarks>
        private static bool TryAlignTableRowCellBreaks(TableBlock table, PdfTableStyle style,
            PreparedFlowTableRows rows, int columns, double[] columnWidths, double columnGap,
            int rowIndex, int startLine, double available, double fullFrameHeight, bool keepSharedBoundary) {
            if (keepSharedBoundary) return false;
            List<TableCellLayout> cells = GetTableCellLayouts(table, rowIndex, columns);
            if (cells.Count < 2 || cells.Any(cell => cell.RowSpan > 1 || cell.Viewport != null || cell.TextRotation != 0)) return false;
            TableCellTextLayout[] layouts = rows.Lines[rowIndex];
            double leading = rows.Leadings[rowIndex];
            double rowFloor = GetTableRowInitialTextHeight(table, rowIndex, columns, leading) +
                GetTableRowMaxPaddingTop(table, style, rowIndex, columns) + GetTableRowMaxPaddingBottom(table, style, rowIndex, columns);
            if (rowFloor > available + 0.001D) return false;

            var ends = new int[cells.Count];
            int boundary = startLine;
            bool continues = false;
            for (int index = 0; index < cells.Count; index++) {
                TableCellLayout cell = cells[index];
                TableCellTextLayout layout = layouts[cell.Column];
                int lineCount = layout.Lines.Count;
                if (startLine >= lineCount) {
                    ends[index] = lineCount;
                    continue;
                }
                double innerWidth = GetTableCellWidth(columnWidths, cell.Column, cell.ColumnSpan, columnGap) -
                    GetTableCellPaddingLeft(style, rowIndex, cell.Column) - GetTableCellPaddingRight(style, rowIndex, cell.Column);
                double padding = GetTableCellPaddingTop(style, rowIndex, cell.Column) + GetTableCellPaddingBottom(style, rowIndex, cell.Column);
                int end = startLine;
                for (int candidate = startLine + 1; candidate <= lineCount; candidate++) {
                    double height = MeasureTableCellContentHeight(cell, layout, startLine, candidate - startLine,
                        leading, innerWidth, includeObjects: startLine == 0) + padding;
                    if (height > available + 0.001D) break;
                    end = candidate;
                }
                var single = new[] { cell };
                while (end > startLine && end < lineCount &&
                    !IsTableRowFragmentBoundaryAllowed(layouts, single, startLine, end - startLine, fullFrameHeight)) end--;
                // A cell that would keep nothing here leaves the decision to the shared rules.
                if (end == startLine) return false;
                ends[index] = end;
                continues |= end < lineCount;
                boundary = Math.Max(boundary, end);
            }
            if (!continues) return false;

            TableCellTextLayout[]? aligned = null;
            for (int index = 0; index < cells.Count; index++) {
                TableCellTextLayout layout = layouts[cells[index].Column];
                if (ends[index] >= boundary || ends[index] >= layout.Lines.Count) continue;
                aligned ??= (TableCellTextLayout[])layouts.Clone();
                aligned[cells[index].Column] = InsertTableCellPlaceholderLines(layout, ends[index], boundary - ends[index]);
            }
            if (aligned == null) return false;
            // Prepared rows share wrapped content with fragment previews, so the row gets a new array.
            rows.Lines[rowIndex] = aligned;
            int maxLines = 1;
            foreach (TableCellLayout cell in cells) maxLines = Math.Max(maxLines, aligned[cell.Column].LineCount);
            rows.LineCounts[rowIndex] = maxLines;
            return true;
        }

        private static TableCellTextLayout InsertTableCellPlaceholderLines(TableCellTextLayout layout, int at, int count) {
            List<T> Insert<T>(List<T> values, T empty) {
                var result = new List<T>(values.Count + count);
                for (int index = 0; index < values.Count; index++) {
                    if (index == at) result.AddRange(Enumerable.Repeat(empty, count));
                    result.Add(values[index]);
                }
                if (at >= values.Count) result.AddRange(Enumerable.Repeat(empty, count));
                return result;
            }
            List<T>? InsertOptional<T>(List<T>? values, T empty) => values == null ? null : Insert(values, empty);

            var lines = new List<List<RichSeg>>(layout.Lines.Count + count);
            for (int index = 0; index < layout.Lines.Count; index++) {
                if (index == at) for (int placeholder = 0; placeholder < count; placeholder++) lines.Add(new List<RichSeg>());
                lines.Add(layout.Lines[index]);
            }
            // A paragraph split at the break owns its placeholders; later paragraphs move after them.
            List<TableCellParagraphRange>? ranges = layout.ParagraphRanges?.Select(range =>
                range.StartLine >= at ? range with { StartLine = range.StartLine + count }
                : range.StartLine + range.LineCount > at ? range with { LineCount = range.LineCount + count }
                : range).ToList();
            List<double> heights = Insert(layout.LineHeights, 0D);
            return new TableCellTextLayout(lines, heights, InsertOptional(layout.LineAlignments, (OfficeAlignment?)null),
                InsertOptional(layout.LineXOffsets, 0D), InsertOptional(layout.LineWidths, 0D), layout.TopSpacing,
                ReferenceEquals(layout.LineBoxHeights, layout.LineHeights) ? heights : Insert(layout.LineBoxHeights, 0D), ranges);
        }
    }
}

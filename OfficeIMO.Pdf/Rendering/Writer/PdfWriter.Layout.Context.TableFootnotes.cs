namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private static bool TableCellHasFootnotes(TableCellLayout cell) =>
            cell.Runs.Any(run => run.FootnoteReference != null) ||
            cell.Paragraphs.Any(paragraph => paragraph.Runs.Any(run => run.FootnoteReference != null));

        /// <summary>Rejects placements that do not yet have a qualified page-associated reference contract.</summary>
        private static bool ValidateTableFootnotePlacement(TableBlock table, PdfTableStyle style, int columns) {
            bool found = false;
            int repeatedHeaders = GetTableRepeatHeaderRowCount(style);
            for (int row = 0; row < table.Rows.Count; row++) {
                foreach (TableCellLayout cell in GetTableCellLayouts(table, row, columns)) {
                    if (!TableCellHasFootnotes(cell)) continue;
                    found = true;
                    if (style.Position != null || !style.ConsumesVerticalFlow)
                        throw new ArgumentException("Footnote references require an ordinary table in vertical document flow.");
                    if (row < repeatedHeaders)
                        throw new ArgumentException("Footnote references in repeating table headers are not supported.");
                    if (cell.TextRotation != 0)
                        throw new ArgumentException("Footnote references in rotated table cells are not supported.");
                }
            }
            return found;
        }

        /// <summary>Collects references only from the same cell lines and clip that the admitted fragment paints.</summary>
        private IEnumerable<RichSeg> GetFlowTableCellFootnoteSegments(PdfTableStyle style, TableCellLayout cell,
            TableCellTextLayout lines, int row, int startLine, int requestedLines, double x, double top,
            double width, double height, double size, double leading, bool bold, double? verticalOffset = null) {
            if (!TableCellHasFootnotes(cell)) return Enumerable.Empty<RichSeg>();
            FlowTableCellTextFrame frame = PrepareFlowTableCellTextFrame(style, cell, lines, row, startLine,
                requestedLines, x, top, width, height, size, leading, bold, verticalOffset);
            var visible = SliceTableCellLines(lines, startLine, frame.VisibleLineCount);
            if (cell.Viewport != null || style.PreservePartialCellLines) {
                var heights = SliceTableCellLineHeights(lines, startLine, frame.VisibleLineCount, leading);
                var alignments = SliceTableCellLineAlignments(lines, startLine, frame.VisibleLineCount);
                var offsets = SliceTableCellLineXOffsets(lines, startLine, frame.VisibleLineCount);
                var widths = SliceTableCellLineWidths(lines, startLine, frame.VisibleLineCount, frame.InnerWidth);
                TableCellContentFrame clip = GetTableCellInkClipFrame(style, row, cell.Column, x, top, width, height);
                PdfAlign alignment = MapTableCellAlignment(GetTableCellAlignment(style, row, cell.Column, cell.Text));
                OmitInvisibleTableCellViewportLines(visible, heights, alignments, offsets, widths, alignment,
                    frame.FirstBaseline, frame.Content.Left + GetTableCellPaddingLeft(style, row, cell.Column),
                    frame.InnerWidth, clip.Left, clip.Top - clip.Height, clip.Width, clip.Height,
                    leading, size, currentOpts, GetTableRowFont(currentOpts, bold));
            }
            return visible.SelectMany(line => line).Where(segment => segment.FootnoteReference != null);
        }

        /// <summary>Previews a prospective row without advancing merged-cell cursors or scheduling any reference.</summary>
        private IEnumerable<RichSeg> GetTableRowFootnoteSegments(TableBlock table, PdfTableStyle style,
            PreparedFlowTableRows rows, int columns, double[] widths, double columnGap, double rowGap,
            TableSpanFlow spans, int row, int startLine, int lineCount, double top, double height,
            double tableX, double fullFrameHeight, double closingPadding) {
            var fragments = new List<(int Row, int Column, IEnumerable<RichSeg> Segments)>();
            bool whole = startLine == 0 && lineCount == rows.LineCounts[row];
            foreach (TableCellLayout cell in GetTableCellLayouts(table, row, columns)) {
                if (spans.Contains(row, cell.Column) || !TableCellHasFootnotes(cell)) continue;
                double cellX = tableX;
                for (int column = 0; column < cell.Column; column++) cellX += widths[column] + columnGap;
                double cellWidth = GetTableCellWidth(widths, cell.Column, cell.ColumnSpan, columnGap);
                double cellHeight = whole && cell.RowSpan > 1 ? GetTableCellHeight(rows.Heights, row, cell.RowSpan, rowGap) : height;
                bool wholeCell = cell.Viewport != null || whole && cell.RowSpan > 1;
                int sourceStart = wholeCell ? 0 : startLine;
                TableCellTextLayout text = rows.Lines[row][cell.Column];
                fragments.Add((row, cell.Column, GetFlowTableCellFootnoteSegments(style, cell, text, row,
                    sourceStart, wholeCell ? text.LineCount : lineCount, cellX, top, cellWidth, cellHeight,
                    rows.Sizes[row], rows.Leadings[row], rows.Bold[row])));
            }
            foreach (TableSpanCellFlow span in spans.Cells) {
                if (span.Row > row || span.EndRow <= row || !TableCellHasFootnotes(span.Cell)) continue;
                double fragmentTop = span.Height > .001D ? span.Top : top;
                double fragmentHeight = fragmentTop - (top - height);
                double processedHeight = span.ProcessedHeight;
                if (span.Height <= .001D && processedHeight > 0D && span.LastFragmentRow < row)
                    processedHeight += (row - span.LastFragmentRow) * rowGap;
                double alignment = ResolveTableSpanAlignmentLeadingHeight(span, style, rows, widths, columnGap, rowGap);
                int admitted = GetTableSpanAdmittedLineCount(span, style, rows, fullFrameHeight,
                    span.EndRow == table.Rows.Count ? closingPadding : 0D, fragmentHeight, processedHeight, alignment);
                double cellX = tableX;
                for (int column = 0; column < span.Cell.Column; column++) cellX += widths[column] + columnGap;
                double cellWidth = GetTableCellWidth(widths, span.Cell.Column, span.Cell.ColumnSpan, columnGap);
                fragments.Add((span.Row, span.Cell.Column, GetFlowTableCellFootnoteSegments(style, span.Cell,
                    rows.Lines[span.Row][span.Cell.Column], span.Row, span.ConsumedLines, admitted,
                    cellX, fragmentTop, cellWidth, fragmentHeight, rows.Sizes[span.Row], rows.Leadings[span.Row],
                    rows.Bold[span.Row], Math.Max(0D, alignment - processedHeight))));
            }
            return fragments.OrderBy(fragment => fragment.Row).ThenBy(fragment => fragment.Column)
                .SelectMany(fragment => fragment.Segments);
        }

        /// <summary>Collects a guaranteed whole row group as one atomic admission, retaining source order.</summary>
        private IEnumerable<RichSeg> GetTableGroupFootnoteSegments(TableBlock table, PdfTableStyle style,
            PreparedFlowTableRows rows, int columns, double[] widths, double columnGap, double rowGap,
            TableSpanFlow spans, int firstRow, int count, double top, double tableX,
            double fullFrameHeight, double closingPadding) {
            var fragments = new List<(int Row, int Column, IEnumerable<RichSeg> Segments)>();
            double rowTop = top;
            int endRow = Math.Min(table.Rows.Count, firstRow + count);
            var rowTops = new Dictionary<int, double>();
            for (int row = firstRow; row < endRow; row++) {
                rowTops[row] = rowTop;
                foreach (TableCellLayout cell in GetTableCellLayouts(table, row, columns)) {
                    if (spans.Contains(row, cell.Column) || !TableCellHasFootnotes(cell)) continue;
                    double cellX = tableX;
                    for (int column = 0; column < cell.Column; column++) cellX += widths[column] + columnGap;
                    double width = GetTableCellWidth(widths, cell.Column, cell.ColumnSpan, columnGap);
                    double height = GetTableCellHeight(rows.Heights, row, Math.Min(cell.RowSpan, endRow - row), rowGap);
                    TableCellTextLayout lines = rows.Lines[row][cell.Column];
                    fragments.Add((row, cell.Column, GetFlowTableCellFootnoteSegments(style, cell, lines, row,
                        0, lines.LineCount, cellX, rowTop, width, height, rows.Sizes[row], rows.Leadings[row], rows.Bold[row])));
                }
                rowTop -= rows.Heights[row] + GetTableRowGapAfter(row, table.Rows.Count, rowGap);
            }
            if (endRow > firstRow) {
                double groupBottom = rowTop + GetTableRowGapAfter(endRow - 1, table.Rows.Count, rowGap);
                foreach (TableSpanCellFlow span in spans.Cells) {
                    if (span.Row >= endRow || span.EndRow <= firstRow || !TableCellHasFootnotes(span.Cell)) continue;
                    int firstPhysicalRow = Math.Max(firstRow, span.Row);
                    double fragmentTop = span.Height > .001D ? span.Top : rowTops[firstPhysicalRow];
                    double spanBottom = span.EndRow < endRow
                        ? rowTops[span.EndRow] + GetTableRowGapAfter(span.EndRow - 1, table.Rows.Count, rowGap)
                        : groupBottom;
                    double fragmentHeight = fragmentTop - spanBottom;
                    double processedHeight = span.ProcessedHeight;
                    if (span.Height <= .001D && processedHeight > 0D && span.LastFragmentRow < firstPhysicalRow)
                        processedHeight += (firstPhysicalRow - span.LastFragmentRow) * rowGap;
                    double alignment = ResolveTableSpanAlignmentLeadingHeight(span, style, rows, widths, columnGap, rowGap);
                    int admitted = GetTableSpanAdmittedLineCount(span, style, rows, fullFrameHeight,
                        span.EndRow == table.Rows.Count ? closingPadding : 0D, fragmentHeight, processedHeight, alignment);
                    double cellX = tableX;
                    for (int column = 0; column < span.Cell.Column; column++) cellX += widths[column] + columnGap;
                    double cellWidth = GetTableCellWidth(widths, span.Cell.Column, span.Cell.ColumnSpan, columnGap);
                    fragments.Add((span.Row, span.Cell.Column, GetFlowTableCellFootnoteSegments(style, span.Cell,
                        rows.Lines[span.Row][span.Cell.Column], span.Row, span.ConsumedLines, admitted,
                        cellX, fragmentTop, cellWidth, fragmentHeight, rows.Sizes[span.Row], rows.Leadings[span.Row],
                        rows.Bold[span.Row], Math.Max(0D, alignment - processedHeight))));
                }
            }
            return fragments.OrderBy(fragment => fragment.Row).ThenBy(fragment => fragment.Column)
                .SelectMany(fragment => fragment.Segments);
        }

        /// <summary>Offers measured merged-cell line boundaries when a note reduces a flexible physical row's capacity.</summary>
        private static IEnumerable<double> GetTableSpanFragmentHeightCandidates(TableSpanFlow spans, PdfTableStyle style,
            PreparedFlowTableRows rows, double[] widths, double columnGap, double rowGap, int row,
            double maximumHeight, double minimumHeight) {
            var heights = new SortedSet<double> { maximumHeight, minimumHeight };
            foreach (TableSpanCellFlow span in spans.Cells) {
                if (span.Row > row || span.EndRow <= row) continue;
                TableCellTextLayout lines = rows.Lines[span.Row][span.Cell.Column];
                double alignment = ResolveTableSpanAlignmentLeadingHeight(span, style, rows, widths, columnGap, rowGap);
                double leadingHeight = Math.Max(0D, alignment - span.ProcessedHeight);
                double padding = GetTableCellPaddingTop(style, span.Row, span.Cell.Column) +
                    GetTableCellPaddingBottom(style, span.Row, span.Cell.Column);
                for (int count = 1; count <= lines.LineCount - span.ConsumedLines; count++) {
                    double height = MeasureTableCellTextHeight(lines, span.ConsumedLines, count, rows.Leadings[span.Row]) +
                        padding + leadingHeight - span.Height;
                    if (height > maximumHeight + .001D) break;
                    if (height >= minimumHeight - .001D) heights.Add(Math.Max(minimumHeight, height));
                }
            }
            return heights.Where(height => height > .001D && height >= minimumHeight - .001D && height <= maximumHeight + .001D)
                .Reverse();
        }
    }
}

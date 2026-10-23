namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Logical row and independently consumed merged-cell content at a physical frame boundary.</summary>
        private sealed class TableFlowCursor {
            public TableFlowCursor(int row, int line, TableSpanFlow spans, bool spanTail = false) {
                Row = row; Line = line; Spans = spans; SpanTail = spanTail;
            }
            public int Row { get; }
            public int Line { get; }
            public TableSpanFlow Spans { get; }
            public bool SpanTail { get; }
        }

        /// <summary>Accepted geometry and span cuts. Painting consumes this plan without repeating admission.</summary>
        private sealed class TableFragmentPlan {
            public TableFragmentPlan(PreparedTableGeometry geometry, TableFlowCursor start, TableFlowCursor next,
                List<TableRowSlice> slices, double before, double after, double height,
                Dictionary<(int Row, int Column), int> finalSpans, double fullHeight) {
                Geometry = geometry; Start = start; Next = next; Slices = slices; Before = before;
                After = after; Height = height; FinalSpans = finalSpans; FullHeight = fullHeight;
            }
            public PreparedTableGeometry Geometry { get; }
            public TableFlowCursor Start { get; }
            public TableFlowCursor Next { get; }
            public List<TableRowSlice> Slices { get; }
            public double Before { get; }
            public double After { get; }
            public double Height { get; }
            public Dictionary<(int Row, int Column), int> FinalSpans { get; }
            public double FullHeight { get; }
            public bool Continues => Next.Row < Geometry.Table.Rows.Count;
        }

        /// <summary>Fits physical rows using the body table's row, paragraph and flexible-span primitives. Preview is inert.</summary>
        private TableFragmentPlan? PlanTableFragment(PreparedTableGeometry prepared, TableFlowCursor? position,
            double available, double fullHeight, bool minimumOnly = false, double finalOutset = 0D) {
            TableBlock table = prepared.Table;
            PdfTableStyle style = position == null ? prepared.Style
                : PrepareTableBorderContinuation(table, prepared.Style, position.Row);
            int columns = prepared.Columns, footerStart = table.Rows.Count - style.FooterRowCount;
            var start = position ?? new TableFlowCursor(0, 0, new TableSpanFlow(table, style, columns, style.HeaderRowCount, footerStart));
            TableSpanFlow spans = start.Spans.CloneCursor();
            PreparedFlowTableRows rows = ReferenceEquals(style, prepared.Style) ? prepared.Rows.WithIndependentHeights()
                : PrepareFlowTableRows(table, style, columns, prepared.Widths, prepared.Gap, prepared.Gap,
                    style.HeaderRowCount, footerStart, prepared.Rows, start.Row, start.Line, spanFlow: spans);
            var geometry = new PreparedTableGeometry(table, style, rows, prepared.Widths, prepared.Width, prepared.CornerRadius);
            var slices = new List<TableRowSlice>();
            int row = start.Row, line = start.Line;
            bool tail = start.SpanTail, initial = position == null;
            double before = initial ? style.SpacingBefore : 0D, used = before;
            double gap = prepared.Gap;
            bool keepWhole = style.KeepTogether && style.SpacingBefore +
                GetTableRowsHeight(prepared.Rows.Heights, 0, table.Rows.Count, gap) + style.SpacingAfter <= fullHeight + .001D;
            if (keepWhole) minimumOnly = false;
            int repeated = !initial && row >= style.HeaderRowCount ? GetTableRepeatHeaderRowCount(style) : 0;
            double repeatedHeight = 0D;
            for (int header = 0; header < repeated; header++)
                repeatedHeight += rows.Heights[header] + GetTableRowGapAfter(header, table.Rows.Count, gap);
            double textFrameHeight = Math.Max(0D, fullHeight - repeatedHeight);
            if (!initial && style.PageContinuationSpacingBefore > 0D) {
                double required = tail ? MeasureTableSpanRemainderHeight(spans, style, rows, prepared.Widths, gap, row,
                    minimumFragmentFrameHeight: textFrameHeight) : MeasurePreparedTableRowSegmentHeight(table, style, rows,
                    columns, prepared.Widths, gap, row, line, 1, false);
                before = Math.Min(style.PageContinuationSpacingBefore, Math.Max(0D, available - repeatedHeight - required));
                used = before;
            }
            for (int header = 0; header < repeated; header++) {
                double headerGap = GetTableRowGapAfter(header, table.Rows.Count, gap);
                slices.Add(new TableRowSlice(header, 0, rows.LineCounts[header], rows.Heights[header], true, true) { GapAfter = headerGap });
                used += rows.Heights[header] + headerGap;
            }
            int contentStart = slices.Count;
            while (row < table.Rows.Count) {
                cancellationToken.ThrowIfCancellationRequested();
                if (slices.Count >= MaximumTextLayoutLines) throw new InvalidDataException("Table flow exceeded the 100,000-fragment limit.");
                double room = available - used;
                if (room <= .001D) break;
                ReserveTableSpanRemainder(spans, style, rows, prepared.Widths, gap, row, textFrameHeight);
                int remaining = rows.LineCounts[row] - line;
                double height;
                int take;
                if (tail) {
                    double required = MeasureTableSpanRemainderHeight(spans, style, rows, prepared.Widths, gap, row);
                    if (row == table.Rows.Count - 1 && required + style.SpacingAfter + finalOutset > room + .001D)
                        room = Math.Max(0D, room - style.SpacingAfter - finalOutset);
                    height = Math.Min(required, room);
                    bool progress = spans.ActiveCells.Any(span => span.EndRow == row + 1 &&
                        GetTableSpanAdmittedLineCount(span, style, rows, textFrameHeight, 0D, height, span.ProcessedHeight,
                            ResolveTableSpanAlignmentLeadingHeight(span, style, rows, prepared.Widths, gap, gap)) > 0);
                    if (!progress) break;
                    take = 0;
                } else {
                    height = MeasurePreparedTableRowSegmentHeight(table, style, rows, columns, prepared.Widths,
                        gap, row, line, remaining, false);
                    if (row == table.Rows.Count - 1 && height + style.SpacingAfter + finalOutset > room + .001D)
                        room = Math.Max(0D, room - style.SpacingAfter - finalOutset);
                    int minimumRows = initial && row == 0
                        ? Math.Min(table.Rows.Count, style.HeaderRowCount + style.MinimumBodyRowsOnFirstPage) : 0;
                    double groupHeight = GetTableRowsHeight(rows.Heights, row, minimumRows, gap);
                    int finalMinimum = Math.Min(style.MinimumBodyRowsOnLastPage, Math.Max(0, footerStart - style.HeaderRowCount));
                    if (line == 0 && finalMinimum > 0 && footerStart - row == finalMinimum) {
                        double finalHeight = GetTableRowsHeight(rows.Heights, row, table.Rows.Count - row, gap) + style.SpacingAfter;
                        if (finalHeight <= textFrameHeight + .001D && finalHeight > room + .001D) break;
                    }
                    if (minimumRows > 0 && groupHeight <= textFrameHeight + .001D && groupHeight > room + .001D) break;
                    take = remaining;
                    bool flexibleSpan = line == 0 && spans.CoversRow(row) &&
                        !GetTableRowFixedHeight(style, row).HasValue;
                    if (flexibleSpan && rows.IntrinsicHeights[row] <= room + .001D &&
                        (height > room + .001D || minimumOnly)) {
                        height = minimumOnly ? rows.IntrinsicHeights[row] : Math.Min(height, room);
                    }
                    bool canSplit = remaining > 1 && GetTableRowAllowBreakAcrossPages(style, row) &&
                        (row >= style.HeaderRowCount || height > textFrameHeight + .001D);
                    if (height > room + .001D || minimumOnly && canSplit) {
                        if (!canSplit) break;
                        take = GetPreparedTableRowSegmentLineCountThatFits(table, style, rows, columns, prepared.Widths,
                            gap, row, line, room, textFrameHeight, canMoveToNextFrame: room + .001D < textFrameHeight,
                            requireDefaultFirstFragment: line == 0);
                        if (minimumOnly && take > 0) {
                            var cells = GetTableCellLayouts(table, row, columns);
                            for (int count = 1; count < take; count++) {
                                if (!IsTableRowFragmentBoundaryAllowed(rows.Lines[row], cells, line, count,
                                    textFrameHeight, requireDefaultFirstFragment: line == 0)) continue;
                                take = count;
                                break;
                            }
                        }
                        if (take == 0) break;
                        height = MeasurePreparedTableRowSegmentHeight(table, style, rows, columns, prepared.Widths,
                            gap, row, line, take, false);
                    }
                }
                var slice = new TableRowSlice(row, line, take, height, row < style.HeaderRowCount && line == 0 && !tail);
                spans.AdmitRow(row, -used, -used - height, gap);
                used += height;
                slices.Add(slice);
                bool rowComplete = tail || line + take == rows.LineCounts[row];
                if (rowComplete) slice.CompletedSpans = AdvancePlannedTableSpans(spans, style, rows, prepared.Widths,
                    gap, textFrameHeight, row);
                bool hasTail = rowComplete && GetTableSpanRemainingLineCount(spans, rows, row) > 0;
                if (hasTail) { line = rows.LineCounts[row]; tail = true; break; }
                if (rowComplete) {
                    slice.GapAfter = GetTableRowGapAfter(row, table.Rows.Count, gap);
                    used += slice.GapAfter;
                    row++; line = 0; tail = false;
                } else { line += take; break; }
                int requiredFirstRows = style.HeaderRowCount + Math.Min(style.MinimumBodyRowsOnFirstPage,
                    Math.Max(0, footerStart - style.HeaderRowCount));
                if (minimumOnly && (!initial || row >= Math.Max(1, requiredFirstRows))) break;
            }
            if (slices.Count == contentStart) return null;
            var finalSpans = AdvancePlannedTableSpans(spans, style, rows, prepared.Widths, gap, textFrameHeight, -1);
            if (row < table.Rows.Count && tail && GetTableSpanRemainingLineCount(spans, rows, row) == 0) {
                row++; line = 0; tail = false;
            }
            double after = row == table.Rows.Count ? style.SpacingAfter : 0D;
            if (used + after + (row == table.Rows.Count ? finalOutset : 0D) > available + .001D) return null;
            if (keepWhole && row < table.Rows.Count) return null;
            return new TableFragmentPlan(geometry, start, new TableFlowCursor(row, line, spans, tail), slices,
                before, after, used + after, finalSpans, textFrameHeight);
        }

        private static Dictionary<(int Row, int Column), int> AdvancePlannedTableSpans(TableSpanFlow spans,
            PdfTableStyle style, PreparedFlowTableRows rows, double[] widths, double gap, double fullHeight, int completedRow) {
            var cuts = new Dictionary<(int Row, int Column), int>();
            foreach (TableSpanCellFlow span in spans.ActiveCells) {
                if (span.Height <= .001D || completedRow >= 0 && span.EndRow != completedRow + 1) continue;
                span.AlignmentLeadingHeight ??= ResolveTableSpanAlignmentLeadingHeight(span, style, rows, widths, gap, gap);
                int count = GetTableSpanAdmittedLineCount(span, style, rows, fullHeight, 0D);
                cuts.Add((span.Row, span.Cell.Column), count);
                span.ConsumedLines += count;
                span.DestinationEmitted = true;
                span.ProcessedHeight += span.Height;
                span.Height = 0D;
            }
            return cuts;
        }

        private void PaintTableFragment(TableFragmentPlan plan, double frameX, double frameWidth, ref double top) {
            PreparedTableGeometry geometry = plan.Geometry;
            TableBlock table = geometry.Table;
            PdfTableStyle style = geometry.Style;
            TableSpanFlow spans = plan.Start.Spans.CloneCursor();
            double left = ResolveTableX(table.Align, style, currentOpts.MarginLeft + frameX, frameWidth, geometry.Width);
            int? noteStructure = RegisterStructureContainer("Note");
            int? structure = RegisterStructureContainer("Table", noteStructure, alternativeText: style.AlternativeText);
            StringBuilder? paired = null;
            top -= plan.Before;
            foreach (TableRowSlice slice in plan.Slices) {
                cancellationToken.ThrowIfCancellationRequested();
                PaintPreparedTableRow(geometry, spans, slice, left, top, structure, ref paired);
                if (slice.CompletedSpans != null)
                    FlushTableSpanFragments(table, style, spans, geometry.Rows, geometry.Widths, geometry.Gap, geometry.Gap,
                        left, structure, ref paired, slice.Row, geometry.CornerRadius, plan.FullHeight,
                        acceptedLineCounts: slice.CompletedSpans);
                top -= slice.Height + slice.GapAfter;
            }
            FlushTableSpanFragments(table, style, spans, geometry.Rows, geometry.Widths, geometry.Gap, geometry.Gap,
                left, structure, ref paired, cornerRadius: geometry.CornerRadius, fullFrameHeight: plan.FullHeight,
                acceptedLineCounts: plan.FinalSpans);
            if (paired != null) sb.Append(paired);
            top -= plan.After;
        }
    }
}

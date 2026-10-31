namespace OfficeIMO.Pdf;

internal static partial class TableDetector {
    /// <summary>
    /// Recovers labelled tables directly from complete ordinary stroke grids when
    /// text alignment found no table. The existing collector owns visibility and
    /// paint admission; merged, competing or incompletely bounded cells decline.
    /// </summary>
    internal static List<PdfStructuredTable> RecoverClosedRulingTables(
        IReadOnlyList<TextLayoutEngine.TextLine> lines, IReadOnlyList<PdfTableRuling> rulings,
        PdfUnderstandingPageContext context) {
        var result = new List<PdfStructuredTable>();
        List<PdfTableRuling> horizontal = MergeRulingSegments(rulings, true, context);
        List<PdfTableRuling> vertical = MergeRulingSegments(rulings, false, context);
        var groups = new List<List<PdfTableRuling>>();
        foreach (PdfTableRuling line in horizontal) {
            context.ConsumeWork();
            List<PdfTableRuling>? group = null;
            foreach (List<PdfTableRuling> candidate in groups) {
                context.ConsumeWork();
                if (Near(line.Left, candidate[0].Left) && Near(line.Right, candidate[0].Right)) {
                    group = candidate;
                    break;
                }
            }
            if (group == null) { group = new List<PdfTableRuling>(); groups.Add(group); }
            group.Add(line);
        }
        foreach (List<PdfTableRuling> group in groups) {
            if (group.Count < 3) continue;
            double left = group[0].Left, right = group[0].Right;
            foreach (PdfTableRuling edge in vertical) {
                context.ConsumeWork();
                if (!Near(edge.Left, left)) continue;
                bool closedRight = false;
                foreach (PdfTableRuling other in vertical) {
                    context.ConsumeWork();
                    if (Near(other.Left, right) && Near(other.Top, edge.Top) && Near(other.Bottom, edge.Bottom)) {
                        closedRight = true;
                        break;
                    }
                }
                if (!closedRight) continue;
                var columns = new List<double>();
                foreach (PdfTableRuling line in vertical) {
                    context.ConsumeWork();
                    if (Near(line.Top, edge.Top) && Near(line.Bottom, edge.Bottom) &&
                        line.Left >= left - RulingTolerance && line.Left <= right + RulingTolerance) columns.Add(line.Left);
                }
                double[] x = DistinctRulingCoordinates(columns, context);
                var rows = new List<double>();
                foreach (PdfTableRuling line in group) {
                    context.ConsumeWork();
                    if (line.Top <= edge.Top + RulingTolerance && line.Top >= edge.Bottom - RulingTolerance) rows.Add(line.Top);
                }
                double[] y = DistinctRulingCoordinates(rows, context);
                Array.Reverse(y);
                if (x.Length < 3 || x.Length > MaximumPositionedRecoveryColumns + 1 || y.Length < 3 ||
                    (long)(x.Length - 1) * (y.Length - 1) > MaximumPositionedRecoveryCells ||
                    !Near(y[0], edge.Top) || !Near(y[y.Length - 1], edge.Bottom) ||
                    !HasOnlyGridRulings(rulings, x, y, context)) continue;
                PdfStructuredTable? table = BuildClosedRulingTable(lines, x, y, context);
                if (table != null) result.Add(table);
            }
        }
        return PdfAdvancedUnderstandingStages.CopyAndSort(context, result, static (first, second) => {
            int top = second.YTop.CompareTo(first.YTop);
            return top != 0 ? top : first.Columns[0].From.CompareTo(second.Columns[0].From);
        }).ToList();
    }

    private static List<PdfTableRuling> MergeRulingSegments(IReadOnlyList<PdfTableRuling> rulings,
        bool horizontal, PdfUnderstandingPageContext context) {
        var input = new List<PdfTableRuling>();
        foreach (PdfTableRuling line in rulings) {
            context.ConsumeWork();
            if (line.Horizontal == horizontal) input.Add(line);
        }
        PdfTableRuling[] sorted = PdfAdvancedUnderstandingStages.CopyAndSort(context, input, (a, b) => {
            int axis = (horizontal ? a.Top : a.Left).CompareTo(horizontal ? b.Top : b.Left);
            return axis != 0 ? axis : (horizontal ? a.Left : a.Bottom).CompareTo(horizontal ? b.Left : b.Bottom);
        });
        var result = new List<PdfTableRuling>();
        if (sorted.Length == 0) return result;
        for (int groupStart = 0; groupStart < sorted.Length;) {
            double coordinate = horizontal ? sorted[groupStart].Top : sorted[groupStart].Left;
            int groupEnd = groupStart + 1;
            while (groupEnd < sorted.Length && Near(horizontal ? sorted[groupEnd].Top : sorted[groupEnd].Left, coordinate)) {
                context.ConsumeWork();
                groupEnd++;
            }
            // Nearby axes can arrive in the opposite interval order. Sort within
            // one axis group before joining; never bridge a disconnected interval.
            PdfTableRuling[] intervals = PdfAdvancedUnderstandingStages.CopyAndSort(context,
                sorted.Skip(groupStart).Take(groupEnd - groupStart).ToArray(),
                (a, b) => (horizontal ? a.Left : a.Bottom).CompareTo(horizontal ? b.Left : b.Bottom));
            double start = horizontal ? intervals[0].Left : intervals[0].Bottom;
            double end = horizontal ? intervals[0].Right : intervals[0].Top;
            for (int index = 1; index < intervals.Length; index++) {
                context.ConsumeWork();
                PdfTableRuling line = intervals[index];
                double from = horizontal ? line.Left : line.Bottom, to = horizontal ? line.Right : line.Top;
                if (from <= end + RulingTolerance) end = Math.Max(end, to);
                else {
                    Add();
                    start = from; end = to;
                }
            }
            Add();
            groupStart = groupEnd;

            void Add() => result.Add(horizontal ? new PdfTableRuling(start, coordinate, end, coordinate)
                : new PdfTableRuling(coordinate, start, coordinate, end));
        }
        return result;
    }

    private static PdfStructuredTable? BuildClosedRulingTable(IReadOnlyList<TextLayoutEngine.TextLine> lines,
        double[] x, double[] y, PdfUnderstandingPageContext context) {
        int columnCount = x.Length - 1, rowCount = y.Length - 1;
        var cells = new List<PdfTextSpan>?[rowCount * columnCount];
        var seen = new HashSet<PdfTextSpan>();
        foreach (TextLayoutEngine.TextLine line in lines) {
            foreach (PdfTextSpan span in line.Spans) {
                context.ConsumeWork();
                if (!seen.Add(span) || span.IsArtifactContent || string.IsNullOrWhiteSpace(span.Text)) continue;
                double end = span.X + Math.Max(0D, span.Advance);
                if (end < x[0] || span.X > x[x.Length - 1] || span.Y > y[0] || span.Y < y[y.Length - 1]) continue;
                if (!span.CanProjectCompleteText(context.Height)) return null;
                int row = -1, column = -1;
                for (int index = 0; index < rowCount; index++) {
                    context.ConsumeWork();
                    if (span.Y < y[index] - RulingTolerance && span.Y > y[index + 1] + RulingTolerance) { row = index; break; }
                }
                for (int index = 0; index < columnCount; index++) {
                    context.ConsumeWork();
                    if (span.X >= x[index] - RulingTolerance && end <= x[index + 1] + RulingTolerance) { column = index; break; }
                }
                // Crossing a ruling is ambiguous; do not manufacture text slices.
                if (row < 0 || column < 0) return null;
                int cell = row * columnCount + column;
                (cells[cell] ??= new List<PdfTextSpan>()).Add(span);
            }
        }
        var table = new PdfStructuredTable { Kind = "closed-ruling-cells", YTop = y[0], YBottom = y[y.Length - 1] };
        for (int column = 0; column < columnCount; column++) table.Columns.Add(new PdfStructuredTableColumn { From = x[column], To = x[column + 1] });
        var sourceLines = new List<TextLayoutEngine.TextLine>();
        var sourceRuns = new List<PdfTextSpan>();
        for (int row = 0; row < rowCount; row++) {
            var values = new string[columnCount];
            var rowRuns = new List<PdfTextSpan>();
            for (int column = 0; column < columnCount; column++) {
                context.ConsumeWork();
                List<PdfTextSpan>? runs = cells[row * columnCount + column];
                if (runs == null) { values[column] = string.Empty; continue; }
                List<TextLayoutEngine.TextLine> cellLines = TextLayoutEngine.BuildLines(runs, context.LayoutOptions.ToEngineOptions(),
                    context.ConsumeWork, context.ThrowIfCancellationRequested);
                values[column] = string.Join(" ", cellLines.Select(static line => line.Text));
                rowRuns.AddRange(runs);
            }
            table.Rows.Add(values);
            sourceLines.Add(new TextLayoutEngine.TextLine((y[row] + y[row + 1]) / 2D, x[0], x[x.Length - 1],
                string.Join(" ", values), rowRuns));
            sourceRuns.AddRange(rowRuns);
        }
        if (sourceRuns.Count == 0 || !LooksLikeHeaderRow(table.Rows[0])) return null;
        table.SourceLines = sourceLines;
        table.SourceRuns = sourceRuns;
        return table;
    }
}

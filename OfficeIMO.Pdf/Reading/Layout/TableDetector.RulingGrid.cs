namespace OfficeIMO.Pdf;

/// <summary>A visible straight stroke centre in the current reading frame.</summary>
internal readonly struct PdfTableRuling {
    internal PdfTableRuling(double x1, double y1, double x2, double y2) {
        Left = Math.Min(x1, x2); Right = Math.Max(x1, x2);
        Bottom = Math.Min(y1, y2); Top = Math.Max(y1, y2);
    }
    internal double Left { get; }
    internal double Right { get; }
    internal double Bottom { get; }
    internal double Top { get; }
    internal bool Horizontal => Top - Bottom <= TableDetector.RulingTolerance;
}

internal static partial class TableDetector {
    // Stroke centres, rather than painted outer edges, define the cell intervals.
    // This small tolerance accommodates numeric serialization without bridging gaps.
    internal const double RulingTolerance = 0.25D;
    internal const int MaximumRulingSegments = 4096;

    /// <summary>
    /// Joins segments that continue one straight line, as a writer that strokes each cell's edges paints a rule one
    /// cell at a time. A gap or an offset keeps segments apart, so a merged cell still interrupts its grid.
    /// </summary>
    private static List<PdfTableRuling> JoinCollinearRulings(IReadOnlyList<PdfTableRuling> rulings, PdfUnderstandingPageContext context) {
        var joined = new List<PdfTableRuling>(rulings.Count);
        foreach (bool horizontal in new[] { true, false }) {
            PdfTableRuling[] ordered = rulings.Where(ruling => ruling.Horizontal == horizontal)
                .OrderBy(ruling => horizontal ? ruling.Top : ruling.Left)
                .ThenBy(ruling => horizontal ? ruling.Left : ruling.Bottom).ToArray();
            int first = joined.Count;
            foreach (PdfTableRuling ruling in ordered) {
                context.ConsumeWork();
                if (joined.Count > first) {
                    PdfTableRuling last = joined[joined.Count - 1];
                    if (horizontal && Near(last.Top, ruling.Top) && ruling.Left <= last.Right + RulingTolerance) {
                        joined[joined.Count - 1] = new PdfTableRuling(last.Left, last.Top, Math.Max(last.Right, ruling.Right), last.Top);
                        continue;
                    }
                    if (!horizontal && Near(last.Left, ruling.Left) && ruling.Bottom <= last.Top + RulingTolerance) {
                        joined[joined.Count - 1] = new PdfTableRuling(last.Left, last.Bottom, last.Left, Math.Max(last.Top, ruling.Top));
                        continue;
                    }
                }
                joined.Add(ruling);
            }
        }
        return joined;
    }

    /// <summary>
    /// Refines geometry only for an existing text table whose complete cells agree with
    /// one closed grid of full-span straight rulings. Fragmented and merged grids retain
    /// text inference. No text, source ownership, row count or column count is changed.
    /// </summary>
    internal static bool TryRefineClosedRulingGrid(PdfStructuredTable table,
        IReadOnlyList<PdfTableRuling> rulings, PdfUnderstandingPageContext context) {
        if (table.Columns.Count < 2 || table.Columns.Count > MaximumPositionedRecoveryColumns ||
            table.Rows.Count < 2 || table.Rows.Count >= MaximumRulingSegments ||
            table.SourceLines.Count != table.Rows.Count || table.SourceRuns.Count == 0 ||
            string.Equals(table.Kind, "leaders", StringComparison.Ordinal)) return false;

        rulings = JoinCollinearRulings(rulings, context);
        var horizontalGroups = new List<List<PdfTableRuling>>();
        foreach (PdfTableRuling ruling in rulings) {
            context.ConsumeWork();
            if (!ruling.Horizontal) continue;
            List<PdfTableRuling>? group = null;
            foreach (List<PdfTableRuling> existing in horizontalGroups) {
                context.ConsumeWork();
                if (Near(ruling.Left, existing[0].Left) && Near(ruling.Right, existing[0].Right)) {
                    group = existing;
                    break;
                }
            }
            if (group is null) {
                group = new List<PdfTableRuling>();
                horizontalGroups.Add(group);
            }
            group.Add(ruling);
        }

        double[]? acceptedX = null, acceptedY = null;
        foreach (List<PdfTableRuling> group in horizontalGroups) {
            context.ConsumeWork();
            if (group.Count < table.Rows.Count + 1) continue;
            double left = group[0].Left, right = group[0].Right;
            double[] ys = DistinctRulingCoordinates(group.Select(static line => line.Top).ToList(), context);
            // Descending Y is the existing native table's top-to-bottom row order.
            Array.Reverse(ys);
            for (int start = 0; start + table.Rows.Count < ys.Length; start++) {
                context.ConsumeWork();
                double top = ys[start], bottom = ys[start + table.Rows.Count];
                var xs = new List<double>();
                foreach (PdfTableRuling line in rulings) {
                    context.ConsumeWork();
                    if (line.Horizontal || !Near(line.Top, top) || !Near(line.Bottom, bottom) ||
                        line.Left < left - RulingTolerance || line.Right > right + RulingTolerance) continue;
                    xs.Add(line.Left);
                }
                double[] x = DistinctRulingCoordinates(xs, context);
                if (x.Length != table.Columns.Count + 1 || !Near(x[0], left) || !Near(x[x.Length - 1], right)) continue;
                var y = new double[table.Rows.Count + 1];
                Array.Copy(ys, start, y, 0, y.Length);
                if (!HasOnlyGridRulings(rulings, x, y, context) || !CellsAgreeWithGrid(table, x, y, context)) continue;
                // Even two plausible grids are insufficient evidence to choose either.
                if (acceptedX is not null) return false;
                acceptedX = x;
                acceptedY = y;
            }
        }
        if (acceptedX is null || acceptedY is null) return false;
        // Admit the complete update before mutating geometry, so cancellation or
        // budget exhaustion cannot leave a partly refined internal candidate.
        context.ConsumeWork(table.Columns.Count + 1L);
        for (int index = 0; index < table.Columns.Count; index++) {
            table.Columns[index].From = acceptedX[index];
            table.Columns[index].To = acceptedX[index + 1];
        }
        table.YTop = acceptedY[0];
        table.YBottom = acceptedY[acceptedY.Length - 1];
        return true;
    }

    private static double[] DistinctRulingCoordinates(List<double> coordinates, PdfUnderstandingPageContext context) {
        double[] ordered = PdfAdvancedUnderstandingStages.CopyAndSort(context, coordinates,
            static (first, second) => first.CompareTo(second));
        var distinct = new List<double>();
        foreach (double coordinate in ordered) {
            context.ConsumeWork();
            if (distinct.Count == 0 || !Near(coordinate, distinct[distinct.Count - 1])) distinct.Add(coordinate);
        }
        return distinct.ToArray();
    }

    private static bool HasOnlyGridRulings(IReadOnlyList<PdfTableRuling> rulings,
        double[] x, double[] y, PdfUnderstandingPageContext context) {
        foreach (PdfTableRuling line in rulings) {
            context.ConsumeWork();
            if (line.Right < x[0] - RulingTolerance || line.Left > x[x.Length - 1] + RulingTolerance ||
                line.Top < y[y.Length - 1] - RulingTolerance || line.Bottom > y[0] + RulingTolerance) continue;
            // A stroke crossing the region must belong wholly to one grid boundary.
            // This also rejects competing inner grids and incomplete extra separators.
            bool belongs = false;
            foreach (double boundary in line.Horizontal ? y : x) {
                context.ConsumeWork();
                if (Near(line.Horizontal ? line.Top : line.Left, boundary)) { belongs = true; break; }
            }
            if (!belongs || line.Left < x[0] - RulingTolerance || line.Right > x[x.Length - 1] + RulingTolerance ||
                line.Bottom < y[y.Length - 1] - RulingTolerance || line.Top > y[0] + RulingTolerance) return false;
        }
        return true;
    }

    private static bool CellsAgreeWithGrid(PdfStructuredTable table, double[] x, double[] y,
        PdfUnderstandingPageContext context) {
        var splits = x.Skip(1).Take(x.Length - 2).ToList();
        var owned = new HashSet<PdfTextSpan>();
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            context.ConsumeWork();
            TextLayoutEngine.TextLine line = table.SourceLines[rowIndex];
            if (table.Rows[rowIndex].Length != table.Columns.Count) return false;
            foreach (PdfTextSpan run in line.Spans) {
                context.ConsumeWork();
                owned.Add(run);
                // Compare each run's original assignment as well as the composed cell
                // text: repeated equal values must not conceal a changed assignment.
                int column = 0;
                while (column + 1 < table.Columns.Count && run.X >= table.Columns[column + 1].From) {
                    context.ConsumeWork();
                    column++;
                }
                if (!IsFiniteRulingCoordinate(run.X) || !IsFiniteRulingCoordinate(run.Y) ||
                    !IsFiniteRulingCoordinate(run.Advance) || run.Advance < 0D ||
                    run.X < x[column] - RulingTolerance || run.X + run.Advance > x[column + 1] + RulingTolerance ||
                    run.Y >= y[rowIndex] - RulingTolerance || run.Y <= y[rowIndex + 1] + RulingTolerance ||
                    !run.CanProjectCompleteText(context.Height)) return false;
            }
            context.ConsumeWork(line.Spans.Count + table.Columns.Count);
            string[] cells = SplitBySplits(line, splits);
            for (int column = 0; column < cells.Length; column++) {
                if (!string.Equals(cells[column], table.Rows[rowIndex][column], StringComparison.Ordinal)) return false;
            }
        }
        foreach (PdfTextSpan run in table.SourceRuns) {
            context.ConsumeWork();
            if (!owned.Contains(run)) return false;
        }
        foreach (PdfTextSpan run in context.DecodedRuns) {
            context.ConsumeWork();
            if (owned.Contains(run) || run.IsArtifactContent || string.IsNullOrWhiteSpace(run.Text)) continue;
            if (run.X + Math.Max(0D, run.Advance) >= x[0] && run.X <= x[x.Length - 1] &&
                run.Y <= y[0] && run.Y >= y[y.Length - 1]) return false;
        }
        return true;
    }

    private static bool Near(double first, double second) => Math.Abs(first - second) <= RulingTolerance;
    private static bool IsFiniteRulingCoordinate(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}

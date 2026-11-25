namespace OfficeIMO.Pdf;

internal static partial class TableDetector {
    // Rules of one boundary drawn in pieces (per cell, per colour, as thin bars) land within a point of each other.
    private const double LatticeAxisTolerance = 1D;
    // Pieces of one rule drawn per cell meet or overlap; an authored gap of a point or more keeps them apart.
    private const double LatticeJoinGap = 0.5D;
    // A rule ending at a crossing rule's centre stops up to half the crossing stroke short of its far edge.
    private const double LatticeTouchTolerance = 1.5D;
    private const double LatticeMinimumCellSize = 2D;
    private const int MaximumLatticeIntersectionChecks = 1 << 20;

    /// <summary>
    /// Recovers tables from connected networks of visible rules. Every rule in a network bounds the table's
    /// columns or rows, even when it is drawn only in the header: spreadsheet exports often separate body
    /// columns by alignment alone and draw their dividers in the header. Each text run belongs to the cell
    /// its start falls in, so a run that overflows an undivided neighbour stays with its own column.
    /// </summary>
    internal static List<PdfStructuredTable> RecoverRulingLatticeTables(
        IReadOnlyList<TextLayoutEngine.TextLine> lines, IReadOnlyList<PdfTableRuling> rulings,
        PdfUnderstandingPageContext context) {
        var result = new List<PdfStructuredTable>();
        List<PdfTableRuling> horizontal = MergeLatticeRules(rulings, true, context);
        List<PdfTableRuling> vertical = MergeLatticeRules(rulings, false, context);
        if (horizontal.Count < 2 || vertical.Count < 2 ||
            (long)horizontal.Count * vertical.Count > MaximumLatticeIntersectionChecks) return result;

        // Union-find over rules that touch.
        int count = horizontal.Count + vertical.Count;
        var parent = new int[count];
        for (int index = 0; index < count; index++) parent[index] = index;
        context.ConsumeWork((long)horizontal.Count * vertical.Count / 64 + 1);
        for (int h = 0; h < horizontal.Count; h++) {
            PdfTableRuling across = horizontal[h];
            for (int v = 0; v < vertical.Count; v++) {
                PdfTableRuling down = vertical[v];
                if (down.Left < across.Left - LatticeTouchTolerance || down.Left > across.Right + LatticeTouchTolerance ||
                    across.Top < down.Bottom - LatticeTouchTolerance || across.Top > down.Top + LatticeTouchTolerance) continue;
                int a = Find(h), b = Find(horizontal.Count + v);
                if (a != b) parent[a] = b;
            }
        }

        var components = new Dictionary<int, (List<PdfTableRuling> Horizontal, List<PdfTableRuling> Vertical)>();
        for (int index = 0; index < count; index++) {
            context.ConsumeWork();
            int root = Find(index);
            if (!components.TryGetValue(root, out var component)) {
                component = (new List<PdfTableRuling>(), new List<PdfTableRuling>());
                components.Add(root, component);
            }
            if (index < horizontal.Count) component.Horizontal.Add(horizontal[index]);
            else component.Vertical.Add(vertical[index - horizontal.Count]);
        }

        List<PdfTextSpan>? spans = null;
        foreach (var component in components.Values) {
            context.ConsumeWork();
            if (component.Horizontal.Count < 2 || component.Vertical.Count < 2) continue;
            double left = double.MaxValue, right = double.MinValue, bottom = double.MaxValue, top = double.MinValue;
            var xs = new List<double>(component.Vertical.Count + 2);
            var ys = new List<double>(component.Horizontal.Count + 2);
            foreach (PdfTableRuling line in component.Horizontal) {
                left = Math.Min(left, line.Left); right = Math.Max(right, line.Right);
                ys.Add(line.Top);
            }
            foreach (PdfTableRuling line in component.Vertical) {
                bottom = Math.Min(bottom, line.Bottom); top = Math.Max(top, line.Top);
                xs.Add(line.Left);
            }
            foreach (PdfTableRuling line in component.Horizontal) { bottom = Math.Min(bottom, line.Top); top = Math.Max(top, line.Top); }
            foreach (PdfTableRuling line in component.Vertical) { left = Math.Min(left, line.Left); right = Math.Max(right, line.Left); }
            xs.Add(left); xs.Add(right);
            ys.Add(bottom); ys.Add(top);
            double[] x = ClusterBoundaries(xs, context);
            double[] y = ClusterBoundaries(ys, context);
            // A boundary must close at least one whole cell: a short stroke inside a cell is artwork.
            for (int pass = 0; pass < 4; pass++) {
                double[] keptX = KeepClosingBoundaries(x, y, component.Vertical, vertical: true, context);
                double[] keptY = KeepClosingBoundaries(y, keptX, component.Horizontal, vertical: false, context);
                bool stable = keptX.Length == x.Length && keptY.Length == y.Length;
                x = keptX;
                y = keptY;
                if (stable) break;
            }
            Array.Reverse(y);
            if (x.Length < 3 || y.Length < 3 || x.Length > MaximumPositionedRecoveryColumns + 1 ||
                (long)(x.Length - 1) * (y.Length - 1) > MaximumPositionedRecoveryCells) continue;
            spans ??= CollectLatticeSpans(lines, context);
            PdfStructuredTable? table = BuildLatticeTable(spans, x, y, component.Vertical, component.Horizontal, context);
            if (table is not null) result.Add(table);
        }
        return PdfAdvancedUnderstandingStages.CopyAndSort(context, result, static (first, second) => {
            int order = second.YTop.CompareTo(first.YTop);
            return order != 0 ? order : first.Columns[0].From.CompareTo(second.Columns[0].From);
        }).ToList();

        int Find(int node) {
            while (parent[node] != node) {
                parent[node] = parent[parent[node]];
                node = parent[node];
            }
            return node;
        }
    }

    private static List<PdfTextSpan> CollectLatticeSpans(IReadOnlyList<TextLayoutEngine.TextLine> lines, PdfUnderstandingPageContext context) {
        var spans = new List<PdfTextSpan>();
        var seen = new HashSet<PdfTextSpan>();
        foreach (TextLayoutEngine.TextLine line in lines) {
            foreach (PdfTextSpan span in line.Spans) {
                context.ConsumeWork();
                // Producers clip each cell's text to its cell, so a tight clip still leaves the text readable;
                // only invisible runs stay out.
                if (span.IsArtifactContent || string.IsNullOrWhiteSpace(span.Text) || !span.IsVisible || span.Color?.A <= 3 ||
                    !IsFiniteRulingCoordinate(span.X) || !IsFiniteRulingCoordinate(span.Y)) continue;
                if (seen.Add(span)) spans.Add(span);
            }
        }
        return spans;
    }

    /// <summary>Joins the pieces of each straight rule, bridging the small gaps producers leave at joints.</summary>
    private static List<PdfTableRuling> MergeLatticeRules(IReadOnlyList<PdfTableRuling> rulings, bool horizontal,
        PdfUnderstandingPageContext context) {
        var input = new List<PdfTableRuling>();
        foreach (PdfTableRuling line in rulings) {
            context.ConsumeWork();
            if (line.Horizontal == horizontal) input.Add(line);
        }
        PdfTableRuling[] sorted = PdfAdvancedUnderstandingStages.CopyAndSort(context, input, (a, b) =>
            (horizontal ? a.Top : a.Left).CompareTo(horizontal ? b.Top : b.Left));
        var result = new List<PdfTableRuling>();
        var group = new List<(double From, double To)>();
        for (int start = 0; start < sorted.Length;) {
            double first = horizontal ? sorted[start].Top : sorted[start].Left;
            double sum = 0D;
            int end = start;
            group.Clear();
            while (end < sorted.Length && (horizontal ? sorted[end].Top : sorted[end].Left) - first <= LatticeAxisTolerance) {
                context.ConsumeWork();
                PdfTableRuling line = sorted[end];
                sum += horizontal ? line.Top : line.Left;
                group.Add(horizontal ? (line.Left, line.Right) : (line.Bottom, line.Top));
                end++;
            }
            double coordinate = sum / (end - start);
            group.Sort(static (a, b) => a.From.CompareTo(b.From));
            double from = group[0].From, to = group[0].To;
            for (int index = 1; index < group.Count; index++) {
                context.ConsumeWork();
                if (group[index].From <= to + LatticeJoinGap) {
                    to = Math.Max(to, group[index].To);
                    continue;
                }
                Add(from, to);
                from = group[index].From;
                to = group[index].To;
            }
            Add(from, to);
            start = end;

            void Add(double lo, double hi) => result.Add(horizontal
                ? new PdfTableRuling(lo, coordinate, hi, coordinate)
                : new PdfTableRuling(coordinate, lo, coordinate, hi));
        }
        return result;
    }

    /// <summary>
    /// Keeps the outer boundaries and every inner one some rule of that orientation draws across a whole cell
    /// of the crossing boundaries (both arrays ascending).
    /// </summary>
    private static double[] KeepClosingBoundaries(double[] boundaries, double[] crossing, List<PdfTableRuling> rules, bool vertical,
        PdfUnderstandingPageContext context) {
        if (boundaries.Length < 3) return boundaries;
        var kept = new List<double>(boundaries.Length) { boundaries[0] };
        for (int index = 1; index < boundaries.Length - 1; index++) {
            double boundary = boundaries[index];
            bool closes = false;
            foreach (PdfTableRuling rule in rules) {
                context.ConsumeWork();
                if (Math.Abs((vertical ? rule.Left : rule.Top) - boundary) > LatticeMinimumCellSize) continue;
                double from = vertical ? rule.Bottom : rule.Left, to = vertical ? rule.Top : rule.Right;
                for (int cell = 0; cell + 1 < crossing.Length; cell++) {
                    if (crossing[cell] >= from - LatticeTouchTolerance && crossing[cell + 1] <= to + LatticeTouchTolerance) { closes = true; break; }
                }
                if (closes) break;
            }
            if (closes) kept.Add(boundary);
        }
        kept.Add(boundaries[boundaries.Length - 1]);
        return kept.Count == boundaries.Length ? boundaries : kept.ToArray();
    }

    /// <summary>
    /// Where no rule separates two rows within a column, the cell spans both: its text, which a producer
    /// often centres vertically, belongs to the span's top cell.
    /// </summary>
    /// <summary>Splits a run that crosses drawn dividers into its cells, between words only.</summary>
    private static bool TrySplitAtColumns(PdfTextSpan span, double[] x, int row, int columnCount, List<PdfTextSpan>?[] cells,
        Dictionary<PdfTextSpan, PdfTextSpan> fragments) {
        string text = span.Text;
        if (Math.Abs(span.RotationDegrees) > 0.01D || !PdfTextAdvanceProjection.TryGetResolvedBoundaries(span, out double[] boundaries) ||
            boundaries.Length != text.Length + 1) return false;
        int start = 0;
        while (start < text.Length) {
            int column = ColumnOf(start);
            int end = start + 1;
            while (end < text.Length && ColumnOf(end) == column) end++;
            // A divider through the middle of a word does not separate values; only split between words.
            if (end < text.Length && !char.IsWhiteSpace(text[end - 1]) && !char.IsWhiteSpace(text[end])) return false;
            int first = start, last = end - 1;
            while (first <= last && char.IsWhiteSpace(text[first])) first++;
            while (last >= first && char.IsWhiteSpace(text[last])) last--;
            if (first <= last) {
                PdfTextSpan piece = span.WithVisualText(text.Substring(first, last - first + 1)).WithOffset(boundaries[first], 0D);
                fragments[piece] = span;
                (cells[row * columnCount + column] ??= new List<PdfTextSpan>()).Add(piece);
            }
            start = end;
        }
        return true;

        int ColumnOf(int index) {
            double centre = span.X + (boundaries[index] + boundaries[index + 1]) / 2D;
            int column = 0;
            while (column + 1 < columnCount && centre >= x[column + 1]) column++;
            return column;
        }
    }

    /// <summary>
    /// The span of a run's painted characters. Character boundaries give it exactly; without them a run far wider
    /// than its characters could be (word spacing stretching a space the decoder trimmed) keeps its ink at the end.
    /// Boundaries are only looked up for runs whose advance reaches past a column boundary.
    /// </summary>
    private static (double Start, double End, bool Reliable) VisibleExtent(PdfTextSpan span, double[] x) {
        double start = span.X, end = span.X + Math.Max(0D, span.Advance);
        bool spansBoundary = false;
        for (int index = 1; index < x.Length - 1 && !spansBoundary; index++) {
            spansBoundary = x[index] > start + 0.5D && x[index] < end - LatticeAxisTolerance;
        }
        string text = span.Text;
        bool padded = text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1]));
        if (!spansBoundary && !padded) return (start, end, true);
        if (Math.Abs(span.RotationDegrees) <= 0.01D && PdfTextAdvanceProjection.TryGetResolvedBoundaries(span, out double[] boundaries) &&
            boundaries.Length == text.Length + 1) {
            int first = 0, last = text.Length - 1;
            while (first <= last && char.IsWhiteSpace(text[first])) first++;
            while (last >= first && char.IsWhiteSpace(text[last])) last--;
            if (first <= last) return (span.X + boundaries[first], span.X + boundaries[last + 1], true);
        }
        double size = Math.Max(1D, span.FontSize);
        double plausible = text.Trim().Length * size + 2D * size;
        if (end - start > plausible) return (end - plausible + 2D * size, end, false);
        return (start, end, true);
    }

    private static void MergeVerticalSpans(List<PdfTextSpan>?[] cells, double[] x, double[] y, List<PdfTableRuling> horizontals,
        PdfUnderstandingPageContext context) {
        int columnCount = x.Length - 1, rowCount = y.Length - 1;
        for (int column = 0; column < columnCount; column++) {
            int top = 0;
            for (int row = 1; row < rowCount; row++) {
                context.ConsumeWork();
                List<PdfTextSpan>? runs = cells[row * columnCount + column];
                if (Drawn(y[row], x[column], x[column + 1])) { top = row; continue; }
                if (runs is null) continue;
                List<PdfTextSpan>? target = cells[top * columnCount + column];
                // Text in a cell whose top neighbour already has text joins it only as the next line of the
                // same paragraph; text further down is a cell of its own behind a border that does not show
                // (painted in the fill's colour).
                if (target is not null && !ContinuesParagraph(target, runs)) { top = row; continue; }
                if (target is null) cells[top * columnCount + column] = runs;
                else target.AddRange(runs);
                cells[row * columnCount + column] = null;
            }
        }

        // The next line of a paragraph sits a line below in the same font, size and colour.
        static bool ContinuesParagraph(List<PdfTextSpan> above, List<PdfTextSpan> below) {
            PdfTextSpan last = above[0], next = below[0];
            foreach (PdfTextSpan span in above) if (span.Y < last.Y) last = span;
            foreach (PdfTextSpan span in below) if (span.Y > next.Y) next = span;
            double gap = last.Y - next.Y, size = Math.Max(last.FontSize, next.FontSize);
            return size > 0D && gap > 0D && gap <= size * 1.6D && Math.Abs(last.FontSize - next.FontSize) <= 0.5D &&
                string.Equals(last.BaseFont, next.BaseFont, StringComparison.Ordinal) && Nullable.Equals(last.Color, next.Color);
        }

        // A merged cell has no rule across it at all; rules inset from the frame or broken at a joint still
        // cover most of the column.
        bool Drawn(double boundary, double left, double right) {
            double covered = 0D;
            foreach (PdfTableRuling rule in horizontals) {
                context.ConsumeWork();
                if (Math.Abs(rule.Top - boundary) > LatticeMinimumCellSize) continue;
                covered += Math.Max(0D, Math.Min(rule.Right, right) - Math.Max(rule.Left, left));
            }
            return covered >= (right - left) / 2D;
        }
    }


    /// <summary>
    /// Counts the text baselines most of a row's filled cells share. Folded data rows put a line in nearly
    /// every cell on each baseline; wrapped header cells, centred at different heights, do not.
    /// </summary>
    private static int CountSharedBaselines(List<(double Y, int Column)> baselines, int filledColumns) {
        if (filledColumns < 2 || baselines.Count < filledColumns * 3) return 0;
        baselines.Sort(static (a, b) => b.Y.CompareTo(a.Y));
        int required = Math.Max(2, (int)Math.Ceiling(filledColumns * 0.6D));
        var columns = new HashSet<int>();
        int shared = 0;
        for (int start = 0; start < baselines.Count;) {
            int end = start;
            columns.Clear();
            while (end < baselines.Count && baselines[start].Y - baselines[end].Y <= LatticeMinimumCellSize) {
                columns.Add(baselines[end].Column);
                end++;
            }
            if (columns.Count >= required) shared++;
            start = end;
        }
        return shared;
    }

    private static double[] ClusterBoundaries(List<double> values, PdfUnderstandingPageContext context) {
        double[] ordered = PdfAdvancedUnderstandingStages.CopyAndSort(context, values, static (a, b) => a.CompareTo(b));
        var result = new List<double>();
        foreach (double value in ordered) {
            context.ConsumeWork();
            if (result.Count > 0 && value - result[result.Count - 1] < LatticeMinimumCellSize) continue;
            result.Add(value);
        }
        // The outer edge wins over a boundary just inside it, so the frame keeps its full extent.
        if (result.Count > 1 && ordered.Length > 0 && ordered[ordered.Length - 1] - result[result.Count - 1] > 0D) {
            result[result.Count - 1] = ordered[ordered.Length - 1];
        }
        return result.ToArray();
    }

    private static PdfStructuredTable? BuildLatticeTable(List<PdfTextSpan> spans, double[] x, double[] y,
        List<PdfTableRuling> verticals, List<PdfTableRuling> horizontals, PdfUnderstandingPageContext context) {
        int columnCount = x.Length - 1, rowCount = y.Length - 1;
        var cells = new List<PdfTextSpan>?[rowCount * columnCount];
        // Pieces of runs split at a divider map back to their run, which the table owns as a whole.
        var fragments = new Dictionary<PdfTextSpan, PdfTextSpan>();
        int owned = 0;
        foreach (PdfTextSpan span in spans) {
            context.ConsumeWork();
            // The run's first visible character places it in a column; the middle of its text height places it
            // in a row. Word spacing can stretch a leading space across a whole column.
            (double inkStart, double inkEnd, bool reliable) = VisibleExtent(span, x);
            double start = inkStart + 0.5D;
            double middle = span.Y + Math.Max(0D, span.FontSize) * 0.3D;
            if (start < x[0] || start > x[columnCount] || middle > y[0] || middle < y[rowCount]) continue;
            int column = Locate(x, start, ascending: true);
            int row = Locate(y, middle, ascending: false);
            // Text overflowing into an undivided neighbour, as spreadsheets let it, stays in its own cell. A run
            // that crosses a drawn divider (a whole row written as one string) is split at the divider by its
            // character positions; without them its cells are in doubt and the grid declines.
            double end = inkEnd;
            if (span.ClipPath is PdfPageClipPath clip && clip.IsRectangle) end = Math.Min(end, clip.X + clip.Width);
            bool crosses = false;
            for (int boundary = column + 1; reliable && !crosses && boundary < columnCount && x[boundary] < end - LatticeAxisTolerance; boundary++) {
                foreach (PdfTableRuling divider in verticals) {
                    context.ConsumeWork();
                    if (Math.Abs(divider.Left - x[boundary]) <= LatticeMinimumCellSize &&
                        middle >= divider.Bottom - LatticeTouchTolerance && middle <= divider.Top + LatticeTouchTolerance) { crosses = true; break; }
                }
            }
            owned++;
            if (!crosses) {
                (cells[row * columnCount + column] ??= new List<PdfTextSpan>()).Add(span);
                continue;
            }
            if (!TrySplitAtColumns(span, x, row, columnCount, cells, fragments)) return null;
        }
        if (owned == 0) return null;
        MergeVerticalSpans(cells, x, y, horizontals, context);

        var values = new string[rowCount][];
        var baselines = new List<(double Y, int Column)>();
        var rowRuns = new List<PdfTextSpan>[rowCount];
        bool[] usedColumns = new bool[columnCount];
        int fullRows = 0, foldedRows = 0, textRows = 0;
        for (int row = 0; row < rowCount; row++) {
            values[row] = new string[columnCount];
            rowRuns[row] = new List<PdfTextSpan>();
            int filled = 0;
            baselines.Clear();
            for (int column = 0; column < columnCount; column++) {
                context.ConsumeWork();
                List<PdfTextSpan>? runs = cells[row * columnCount + column];
                if (runs is null) { values[row][column] = string.Empty; continue; }
                if (runs.Count == 1) {
                    // Most cells hold one run; it is its own line.
                    values[row][column] = runs[0].Text.Trim();
                    baselines.Add((runs[0].Y, column));
                } else {
                    List<TextLayoutEngine.TextLine> cellLines = TextLayoutEngine.BuildLines(runs, context.LayoutOptions.ToEngineOptions(),
                        context.ConsumeWork, context.ThrowIfCancellationRequested);
                    values[row][column] = string.Join(" ", cellLines.Select(static line => line.Text.Trim()).Where(static text => text.Length > 0));
                    foreach (TextLayoutEngine.TextLine line in cellLines) baselines.Add((line.Y, column));
                }
                foreach (PdfTextSpan run in runs) {
                    PdfTextSpan source = fragments.TryGetValue(run, out PdfTextSpan? whole) ? whole : run;
                    if (!ReferenceEquals(source, run) && rowRuns[row].Contains(source)) continue;
                    rowRuns[row].Add(source);
                }
                usedColumns[column] = true;
                filled++;
            }
            if (filled >= 2) fullRows++;
            if (CountSharedBaselines(baselines, filled) >= 3) foldedRows++;
            if (filled > 0) textRows++;
        }
        // A boxed paragraph, a frame around a page or a form whose boxes each hold one label is not a table:
        // most rows of a table fill two cells or more.
        if (fullRows < 2 || fullRows * 2 < textRows) return null;
        // Rows whose cells hold several lines on shared baselines are rows the producer did not rule (rules only
        // above and below a block of data). When they make up a good part of the table the grid would fold them
        // together; a wrapped header among ruled rows is not that.
        if (foldedRows * 3 >= textRows) return null;

        // Columns with no text anywhere come from doubled borders, decorative edges and stray dividers.
        var keptColumns = new List<int>(columnCount);
        for (int column = 0; column < columnCount; column++) {
            if (usedColumns[column]) keptColumns.Add(column);
        }
        if (keptColumns.Count < 2) return null;
        // Form boxes are sparse and hold labels and instructions; table cells are mostly filled with short values.
        int filledCells = 0;
        long textLength = 0;
        for (int row = 0; row < rowCount; row++) {
            foreach (int column in keptColumns) {
                string value = values[row][column];
                if (value.Length == 0) continue;
                filledCells++;
                textLength += value.Length;
            }
        }
        if (filledCells < (long)textRows * keptColumns.Count * 0.35D || textLength > filledCells * 80L) return null;

        var table = new PdfStructuredTable { Kind = "ruling-lattice", YTop = y[0], YBottom = y[rowCount] };
        for (int index = 0; index < keptColumns.Count; index++) {
            int column = keptColumns[index];
            // A dropped empty column's width goes to its left neighbour so the kept columns still tile the grid.
            double to = index + 1 < keptColumns.Count ? x[keptColumns[index + 1]] : x[columnCount];
            double from = index == 0 ? x[0] : x[column];
            table.Columns.Add(new PdfStructuredTableColumn { From = from, To = to });
        }
        var sourceLines = new List<TextLayoutEngine.TextLine>(rowCount);
        var sourceRuns = new List<PdfTextSpan>(owned);
        for (int row = 0; row < rowCount; row++) {
            context.ConsumeWork();
            bool empty = rowRuns[row].Count == 0;
            // Spacer rows between doubled rules carry nothing.
            if (empty && y[row] - y[row + 1] < 6D) continue;
            var cellsInRow = new string[keptColumns.Count];
            for (int index = 0; index < keptColumns.Count; index++) cellsInRow[index] = values[row][keptColumns[index]];
            table.Rows.Add(cellsInRow);
            sourceLines.Add(new TextLayoutEngine.TextLine((y[row] + y[row + 1]) / 2D, x[0], x[columnCount],
                string.Join(" ", cellsInRow), rowRuns[row]));
            sourceRuns.AddRange(rowRuns[row]);
        }
        if (table.Rows.Count < 2) return null;
        table.SourceLines = sourceLines;
        table.SourceRuns = sourceRuns;
        return table;

        static int Locate(double[] boundaries, double value, bool ascending) {
            int last = boundaries.Length - 2;
            for (int index = 0; index < last; index++) {
                if (ascending ? value < boundaries[index + 1] : value > boundaries[index + 1]) return index;
            }
            return last;
        }
    }
}

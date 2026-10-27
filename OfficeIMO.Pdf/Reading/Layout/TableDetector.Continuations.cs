namespace OfficeIMO.Pdf;

internal static partial class TableDetector {
    /// <summary>
    /// Joins table pieces that the band detector cut apart. Band detection breaks a table wherever a row is
    /// taller than the usual pitch (a wrapped cell) or a row leaves a column empty, which leaves several small
    /// tables with the same columns stacked directly above each other. Pieces join when they share column
    /// geometry, sit within a few row pitches of each other and no other text lies between them.
    /// </summary>
    internal static void MergeStackedTablePieces(
        List<PdfStructuredTable> tables,
        IReadOnlyList<TextLayoutEngine.TextLine> pageLines,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        if (tables.Count < 2) return;
        tables.Sort(static (first, second) => second.YTop.CompareTo(first.YTop));
        for (int index = 0; index + 1 < tables.Count;) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            PdfStructuredTable upper = tables[index];
            PdfStructuredTable lower = tables[index + 1];
            if (!CanMergeStackedPieces(upper, lower, pageLines)) {
                index++;
                continue;
            }
            upper.Rows.AddRange(lower.Rows);
            upper.SourceLines = upper.SourceLines.Concat(lower.SourceLines).ToList();
            upper.SourceRuns = upper.SourceRuns.Concat(lower.SourceRuns).Distinct().ToArray();
            upper.YBottom = lower.YBottom;
            for (int column = 0; column < upper.Columns.Count; column++) {
                upper.Columns[column].From = Math.Min(upper.Columns[column].From, lower.Columns[column].From);
                upper.Columns[column].To = Math.Max(upper.Columns[column].To, lower.Columns[column].To);
                upper.Columns[column].VisualBounds = null;
            }
            upper.VisualBounds = null;
            tables.RemoveAt(index + 1);
        }
    }

    /// <summary>
    /// Adds the label row directly above a table when the band detector left it out, typically because its labels
    /// sit far apart (and so form several lines) or the row is taller than the body rows. The row must lie within a
    /// couple of row pitches above the table, fit its columns and contain labels rather than values.
    /// </summary>
    internal static void AttachDetachedHeaderRows(
        List<PdfStructuredTable> tables,
        IReadOnlyList<TextLayoutEngine.TextLine> pageLines,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        if (tables.Count == 0) return;
        var owned = new HashSet<TextLayoutEngine.TextLine>(tables.SelectMany(static table => table.SourceLines));
        foreach (PdfStructuredTable table in tables) {
            cancellationCheck?.Invoke();
            if (string.Equals(table.Kind, "leaders", StringComparison.Ordinal) || table.Columns.Count < 2 ||
                table.SourceLines.Count < 2 || table.Rows.Count != table.SourceLines.Count) continue;
            double pitch = EstimateRowPitch(table, table);
            // Header labels may start left of right-aligned first-column values ("#" over "1", "2", ...).
            double slack = Math.Max(24D, pitch * 1.5D);
            double left = table.Columns[0].From - slack;
            double right = table.Columns[table.Columns.Count - 1].To + slack;
            double firstRowY = table.SourceLines.Max(static line => line.Y);
            TextLayoutEngine.TextLine[] candidates = pageLines
                .Where(line => !owned.Contains(line) && line.Y > firstRowY + 0.5D && line.Y - firstRowY <= pitch * 2D &&
                    line.XStart >= left && line.XEnd <= right && !string.IsNullOrWhiteSpace(line.Text))
                .ToArray();
            consumeWork?.Invoke(candidates.Length + 1);
            if (candidates.Length == 0) continue;
            double headerY = candidates.Min(static line => line.Y);
            TextLayoutEngine.TextLine[] row = candidates.Where(line => Math.Abs(line.Y - headerY) <= 1.5D).ToArray();
            if (row.Length != candidates.Length) continue;
            TextLayoutEngine.TextLine? header = row.Length == 1 ? row[0] : TryCombineSameBaselineLines(row.ToList());
            if (header is null) continue;
            string[] cells = AssignHeaderSpansToColumns(table, header);
            int filled = cells.Count(static cell => !string.IsNullOrWhiteSpace(cell));
            if (filled < 2 || filled * 2 < cells.Length || !LooksLikeHeaderLabels(cells)) continue;
            table.Rows.Insert(0, cells);
            var sourceLines = new List<TextLayoutEngine.TextLine>(table.SourceLines.Count + 1) { header };
            sourceLines.AddRange(table.SourceLines);
            table.SourceLines = sourceLines;
            table.SourceRuns = header.Spans.Concat(table.SourceRuns).Distinct().ToArray();
            table.YTop = Math.Max(table.YTop, header.Y);
            table.VisualBounds = null;
            foreach (TextLayoutEngine.TextLine line in row) owned.Add(line);
        }
    }

    /// <summary>
    /// Header labels are often centred or right-aligned over their column, so a label can start left of the
    /// boundary between body cells. Each label span goes to the column whose centre is nearest its own centre.
    /// </summary>
    private static string[] AssignHeaderSpansToColumns(PdfStructuredTable table, TextLayoutEngine.TextLine header) {
        var fragments = new List<PdfTextSpan>[table.Columns.Count];
        for (int index = 0; index < fragments.Length; index++) fragments[index] = new List<PdfTextSpan>();
        foreach (PdfTextSpan span in header.Spans) {
            double center = span.X + Math.Max(0D, span.Advance) / 2D;
            int best = 0;
            double bestDistance = double.MaxValue;
            for (int column = 0; column < table.Columns.Count; column++) {
                double columnCenter = (table.Columns[column].From + table.Columns[column].To) / 2D;
                double distance = Math.Abs(center - columnCenter);
                if (distance < bestDistance) {
                    best = column;
                    bestDistance = distance;
                }
            }
            fragments[best].Add(span);
        }
        return fragments.Select(fragment => ComposeCell(fragment, header.ReadingDirection)).ToArray();
    }

    /// <summary>
    /// Extends a table downwards with rows the band detector could not see: rows whose cells wrap onto several
    /// lines, or rows that leave some columns empty. Each following baseline is split at the table's column
    /// boundaries; a line closer to the previous one than a row pitch continues the cells above it.
    /// </summary>
    internal static void AbsorbWrappedRows(
        List<PdfStructuredTable> tables,
        IReadOnlyList<TextLayoutEngine.TextLine> pageLines,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        if (tables.Count == 0) return;
        var owned = new HashSet<TextLayoutEngine.TextLine>(tables.SelectMany(static table => table.SourceLines));
        var absorbedTables = new HashSet<PdfStructuredTable>();
        foreach (PdfStructuredTable table in tables) {
            cancellationCheck?.Invoke();
            if (absorbedTables.Contains(table) ||
                string.Equals(table.Kind, "leaders", StringComparison.Ordinal) || table.Columns.Count < 2 ||
                table.SourceLines.Count < 2 || table.Rows.Count < 2) continue;
            // Fragments found below this table (single wrapped rows) may already be small tables of their own.
            var fragmentLines = new HashSet<TextLayoutEngine.TextLine>(tables
                .Where(other => !ReferenceEquals(other, table) && !absorbedTables.Contains(other) &&
                    !string.Equals(other.Kind, "leaders", StringComparison.Ordinal) && other.YTop < table.YBottom)
                .SelectMany(static other => other.SourceLines));
            double fontSize = table.SourceLines.SelectMany(static line => line.Spans).Select(static span => span.FontSize).DefaultIfEmpty(10D).Max();
            double pitch = Math.Max(EstimateRowPitch(table, table), fontSize * 1.6D);
            double left = table.Columns[0].From - 4D;
            double right = table.Columns[table.Columns.Count - 1].To + Math.Max(24D, fontSize * 2D);
            // A wrapped cell line sits clearly closer to the line above than the table's row pitch.
            double lineGap = Math.Min(fontSize * 1.6D, EstimateRowPitch(table, table) * 0.75D);
            // Wrapped cell text the band detector already folded into a row is owned through its runs.
            var tableRuns = new HashSet<PdfTextSpan>(table.SourceRuns);
            double cursor = Math.Min(table.SourceLines.Min(static line => line.Y),
                table.SourceRuns.Count == 0 ? double.MaxValue : table.SourceRuns.Min(static run => run.Y));
            var lines = new List<TextLayoutEngine.TextLine>(table.SourceLines);
            var consumed = new HashSet<TextLayoutEngine.TextLine>();
            int added = 0;
            while (added < MaximumRulingSegments) {
                consumeWork?.Invoke(1);
                double maximumGap = Math.Max(pitch * 1.8D, fontSize * 4.5D);
                TextLayoutEngine.TextLine[] below = pageLines
                    .Where(line => (!owned.Contains(line) || fragmentLines.Contains(line)) && !line.Spans.All(tableRuns.Contains) &&
                        line.Y < cursor - 0.5D && cursor - line.Y <= maximumGap &&
                        line.XEnd > left && line.XStart < right && !string.IsNullOrWhiteSpace(line.Text))
                    .ToArray();
                if (below.Length == 0) break;
                double rowY = below.Max(static line => line.Y);
                // Text anywhere else on the page between the table and this line (a margin note, a label) ends it.
                bool interrupted = pageLines.Any(line => !owned.Contains(line) && !line.Spans.All(tableRuns.Contains) &&
                    line.Y < cursor - 0.5D && line.Y > rowY + 1.5D && !string.IsNullOrWhiteSpace(line.Text));
                if (interrupted) break;
                TextLayoutEngine.TextLine[] baseline = below.Where(line => Math.Abs(line.Y - rowY) <= 1.5D).ToArray();
                // Text outside the table's horizontal span ends the table.
                if (baseline.Any(line => line.XStart < left || line.XEnd > right)) break;
                TextLayoutEngine.TextLine? row = baseline.Length == 1 ? baseline[0] : TryCombineSameBaselineLines(baseline.ToList());
                if (row is null) break;
                // A list or a heading under the table is the next block, not a row.
                if (ContentStructureExtractor.IsListItemText(row.Text.Trim())) break;
                if (row.Spans.Any(span => span.FontSize > fontSize * 1.15D)) break;
                // An emphasized row under plain rows starts the next table.
                if (lines.Count > 0 && HasEmphasizedText(row) && !HasEmphasizedText(lines[lines.Count - 1])) break;
                string[]? cells = SplitIntoTableColumns(table, row);
                if (cells is null) break;
                int filled = cells.Count(static cell => !string.IsNullOrWhiteSpace(cell));
                if (!AlignsWithTableColumns(table, lines, row, fontSize)) break;
                if (table.Rows.Count > 0 && cells.SequenceEqual(table.Rows[0], StringComparer.Ordinal)) break;
                bool continuation = cursor - rowY <= lineGap;
                if (continuation) {
                    string[] previous = table.Rows[table.Rows.Count - 1];
                    for (int column = 0; column < cells.Length && column < previous.Length; column++) {
                        if (string.IsNullOrWhiteSpace(cells[column])) continue;
                        previous[column] = string.IsNullOrWhiteSpace(previous[column]) ? cells[column] : previous[column] + " " + cells[column];
                    }
                } else {
                    // A row with a single filled cell is more often a label or note below the table than data.
                    if (filled < 2) break;
                    table.Rows.Add(cells);
                }
                foreach (TextLayoutEngine.TextLine line in baseline) {
                    owned.Add(line);
                    consumed.Add(line);
                }
                if (continuation && lines.Count > 0) {
                    // Keep one source line per row: consumers map a source line index to its row.
                    TextLayoutEngine.TextLine previousLine = lines[lines.Count - 1];
                    lines[lines.Count - 1] = new TextLayoutEngine.TextLine(
                        previousLine.Y,
                        Math.Min(previousLine.XStart, row.XStart),
                        Math.Max(previousLine.XEnd, row.XEnd),
                        previousLine.Text + " " + row.Text,
                        previousLine.Spans.Concat(row.Spans).ToList(),
                        previousLine.ReadingDirection);
                } else {
                    lines.Add(row);
                }
                table.SourceRuns = table.SourceRuns.Concat(row.Spans).Distinct().ToArray();
                table.YBottom = Math.Min(table.YBottom, rowY);
                cursor = rowY;
                added++;
            }
            if (added == 0) continue;
            table.SourceLines = lines;
            table.VisualBounds = null;
            var absorbedLines = new HashSet<TextLayoutEngine.TextLine>(lines.Concat(consumed));
            foreach (PdfStructuredTable other in tables) {
                if (!ReferenceEquals(other, table) && other.SourceLines.Count > 0 && other.SourceLines.All(absorbedLines.Contains)) {
                    absorbedTables.Add(other);
                }
            }
        }
        tables.RemoveAll(absorbedTables.Contains);
    }

    /// <summary>
    /// A new or continued row must start its cells where the table's rows already start (left-aligned) or end
    /// them where the rows end (right-aligned); a line that drifts between those edges is not part of the grid.
    /// </summary>
    private static bool AlignsWithTableColumns(PdfStructuredTable table, IReadOnlyList<TextLayoutEngine.TextLine> rows,
        TextLayoutEngine.TextLine candidate, double fontSize) {
        double tolerance = Math.Max(2.5D, fontSize * 0.3D);
        for (int column = 0; column < table.Columns.Count; column++) {
            double from = column == 0 ? double.MinValue : table.Columns[column].From;
            double to = column + 1 == table.Columns.Count ? double.MaxValue : table.Columns[column].To;
            PdfTextSpan[] spans = candidate.Spans.Where(span => span.X >= from && span.X < to).ToArray();
            if (spans.Length == 0) continue;
            double start = spans.Min(static span => span.X);
            double end = spans.Max(static span => span.X + Math.Max(0D, span.Advance));
            var starts = new List<double>();
            var ends = new List<double>();
            foreach (TextLayoutEngine.TextLine row in rows) {
                PdfTextSpan[] existing = row.Spans.Where(span => span.X >= from && span.X < to).ToArray();
                if (existing.Length == 0) continue;
                starts.Add(existing.Min(static span => span.X));
                ends.Add(existing.Max(static span => span.X + Math.Max(0D, span.Advance)));
            }
            if (starts.Count == 0) continue;
            bool leftAligned = starts.Any(value => Math.Abs(value - start) <= tolerance);
            bool rightAligned = ends.Any(value => Math.Abs(value - end) <= tolerance);
            if (!leftAligned && !rightAligned) return false;
        }
        return true;
    }

    /// <summary>Splits a line at the table's column boundaries, or returns null when a span straddles one.</summary>
    private static string[]? SplitIntoTableColumns(PdfStructuredTable table, TextLayoutEngine.TextLine line) {
        var fragments = new List<PdfTextSpan>[table.Columns.Count];
        for (int index = 0; index < fragments.Length; index++) fragments[index] = new List<PdfTextSpan>();
        foreach (PdfTextSpan span in line.Spans) {
            int column = 0;
            while (column + 1 < table.Columns.Count && span.X >= table.Columns[column].To) column++;
            double end = span.X + Math.Max(0D, span.Advance);
            // Measured advances overshoot slightly; only a clear crossing (a paragraph line) is a straddle.
            if (column + 1 < table.Columns.Count && end > table.Columns[column + 1].From + Math.Max(12D, span.FontSize * 1.5D)) return null;
            fragments[column].Add(span);
        }
        return fragments.Select(fragment => ComposeCell(fragment, line.ReadingDirection)).ToArray();
    }

    /// <summary>
    /// Splits a column at a whitespace channel that runs through every row. Right-aligned figures end close to the
    /// next left-aligned column (themed tables pad cells, not gaps), so the gap between them can be narrower than
    /// the generic cell-gap threshold, yet it sits at the same position on every row.
    /// </summary>
    internal static void SplitColumnsAtAlignedChannels(List<PdfStructuredTable> tables, Action? cancellationCheck) {
        foreach (PdfStructuredTable table in tables) {
            cancellationCheck?.Invoke();
            if (string.Equals(table.Kind, "leaders", StringComparison.Ordinal) || table.Columns.Count == 0 ||
                table.SourceLines.Count < 3 || table.Rows.Count != table.SourceLines.Count) continue;
            for (int column = 0; column < table.Columns.Count && table.Columns.Count < MaximumPositionedRecoveryColumns; column++) {
                double? split = FindAlignedChannel(table, column);
                if (!split.HasValue) continue;
                PdfStructuredTableColumn source = table.Columns[column];
                table.Columns.Insert(column + 1, new PdfStructuredTableColumn { From = split.Value, To = source.To });
                source.To = split.Value;
                source.VisualBounds = null;
                for (int row = 0; row < table.Rows.Count; row++) {
                    string[]? cells = SplitIntoTableColumns(table, table.SourceLines[row]);
                    table.Rows[row] = cells ?? InsertEmptyCell(table.Rows[row], column + 1);
                }
                column++;
            }
        }
    }

    private static string[] InsertEmptyCell(string[] row, int index) {
        var cells = row.ToList();
        cells.Insert(Math.Min(index, cells.Count), string.Empty);
        return cells.ToArray();
    }

    private static double? FindAlignedChannel(PdfStructuredTable table, int column) {
        double from = table.Columns[column].From, to = table.Columns[column].To;
        // Start with the whole column as free and remove every span extent row by row.
        var free = new List<(double Start, double End)> { (from, to) };
        int rowsWithText = 0;
        double fontSize = 10D;
        foreach (TextLayoutEngine.TextLine line in table.SourceLines) {
            PdfTextSpan[] spans = line.Spans.Where(span => span.X >= from - 0.5D && span.X < to).ToArray();
            if (spans.Length == 0) continue;
            rowsWithText++;
            fontSize = Math.Max(fontSize, spans.Max(static span => span.FontSize));
            foreach (PdfTextSpan span in spans) {
                double start = span.X, end = span.X + Math.Max(0D, span.Advance);
                var next = new List<(double Start, double End)>();
                foreach ((double freeStart, double freeEnd) in free) {
                    if (end <= freeStart || start >= freeEnd) { next.Add((freeStart, freeEnd)); continue; }
                    if (start > freeStart) next.Add((freeStart, start));
                    if (end < freeEnd) next.Add((end, freeEnd));
                }
                free = next;
            }
        }
        if (rowsWithText < 3) return null;
        double minimum = Math.Max(6D, fontSize * 0.6D);
        foreach ((double start, double end) in free.OrderByDescending(static gap => gap.End - gap.Start)) {
            if (end - start < minimum || start <= from + 0.5D || end >= to - 0.5D) continue;
            double split = (start + end) / 2D;
            int left = 0, right = 0;
            foreach (TextLayoutEngine.TextLine line in table.SourceLines) {
                if (line.Spans.Any(span => span.X >= from - 0.5D && span.X < split)) left++;
                if (line.Spans.Any(span => span.X >= split && span.X < to)) right++;
            }
            // Both sides must carry content on most rows; otherwise this is ragged text, not two columns.
            if (left * 10 >= rowsWithText * 8 && right * 10 >= rowsWithText * 8) return split;
        }
        return null;
    }

    private static bool LooksLikeHeaderLabels(string[] cells) {
        int letters = 0;
        foreach (string raw in cells) {
            string cell = raw?.Trim() ?? string.Empty;
            if (cell.Length == 0) continue;
            if (PdfUnicodeScalarAnalysis.ContainsLetter(cell)) {
                // A label is not a value: no amounts, percentages or dates.
                if (PdfLogicalTableAnalysis.TryParseNumericValue(cell, System.Globalization.CultureInfo.InvariantCulture, out _)) return false;
                letters++;
                continue;
            }
            if (!(cell.Length <= 2 && cell.All(static ch => char.IsSymbol(ch) || char.IsPunctuation(ch)))) return false;
        }
        return letters > 0;
    }

    private static bool CanMergeStackedPieces(
        PdfStructuredTable upper,
        PdfStructuredTable lower,
        IReadOnlyList<TextLayoutEngine.TextLine> pageLines) {
        if (string.Equals(upper.Kind, "leaders", StringComparison.Ordinal) ||
            string.Equals(lower.Kind, "leaders", StringComparison.Ordinal)) return false;
        if (upper.Columns.Count < 2 || upper.Columns.Count != lower.Columns.Count) return false;
        if (upper.SourceLines.Count == 0 || lower.SourceLines.Count == 0) return false;
        if (lower.YTop > upper.YBottom) return false;

        // Internal column boundaries agree within a fraction of the narrower neighbouring column.
        for (int column = 0; column + 1 < upper.Columns.Count; column++) {
            double width = Math.Min(
                Math.Min(upper.Columns[column].To - upper.Columns[column].From, upper.Columns[column + 1].To - upper.Columns[column + 1].From),
                Math.Min(lower.Columns[column].To - lower.Columns[column].From, lower.Columns[column + 1].To - lower.Columns[column + 1].From));
            double tolerance = Math.Max(6D, Math.Max(0D, width) * 0.35D);
            if (Math.Abs(upper.Columns[column].To - lower.Columns[column].To) > tolerance) return false;
        }

        double pitch = EstimateRowPitch(upper, lower);
        double gap = upper.YBottom - lower.YTop;
        if (gap > pitch * 3.5D) return false;

        // Any other text between the pieces (a heading, a label, a margin note) separates two tables, and a piece that
        // starts by repeating the upper piece's first row is a new table with its own header.
        var owned = new HashSet<TextLayoutEngine.TextLine>(upper.SourceLines.Concat(lower.SourceLines));
        for (int index = 0; index < pageLines.Count; index++) {
            TextLayoutEngine.TextLine line = pageLines[index];
            if (owned.Contains(line)) continue;
            if (line.Y >= upper.YBottom - 0.5D || line.Y <= lower.YTop + 0.5D) continue;
            if (string.IsNullOrWhiteSpace(line.Text)) continue;
            return false;
        }
        if (upper.Rows.Count > 0 && lower.Rows.Count > 0 && upper.Rows[0].SequenceEqual(lower.Rows[0], StringComparer.Ordinal)) return false;
        // An emphasized first row under plain rows is the header of the next table.
        TextLayoutEngine.TextLine lowerFirst = lower.SourceLines.OrderByDescending(static line => line.Y).First();
        TextLayoutEngine.TextLine upperLast = upper.SourceLines.OrderBy(static line => line.Y).First();
        if (HasEmphasizedText(lowerFirst) && !HasEmphasizedText(upperLast)) return false;
        return true;
    }

    private static double EstimateRowPitch(PdfStructuredTable upper, PdfStructuredTable lower) {
        var gaps = new List<double>();
        foreach (PdfStructuredTable table in new[] { upper, lower }) {
            double[] ys = table.SourceLines.Select(static line => line.Y).Distinct().OrderByDescending(static y => y).ToArray();
            for (int index = 1; index < ys.Length; index++) {
                double gap = ys[index - 1] - ys[index];
                if (gap > 0.5D) gaps.Add(gap);
            }
        }
        if (gaps.Count == 0) {
            double fontSize = upper.SourceLines.Concat(lower.SourceLines).SelectMany(static line => line.Spans)
                .Select(static span => span.FontSize).DefaultIfEmpty(10D).Max();
            return Math.Max(8D, fontSize * 2D);
        }
        gaps.Sort();
        return gaps[gaps.Count / 2];
    }
}

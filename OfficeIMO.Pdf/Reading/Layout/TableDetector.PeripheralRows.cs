namespace OfficeIMO.Pdf;

internal static partial class TableDetector {
    /// <summary>
    /// Recovers a compact header/body region when an unrelated side annotation changes
    /// the middle row's split count. Only fully contained source runs belong to the table.
    /// </summary>
    private static void RecoverTablesWithPeripheralRowText(List<PdfStructuredTable> tables,
        IReadOnlyList<TextLayoutEngine.TextLine> pageLines, Action<long>? consumeWork, Action? cancellationCheck) {
        var rows = new List<TextLayoutEngine.TextLine>();
        var baseline = new List<TextLayoutEngine.TextLine>();
        foreach (TextLayoutEngine.TextLine line in pageLines.Take(MaximumPositionedRecoveryLines)
            .OrderByDescending(static line => line.Y)) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            if (baseline.Count > 0 && Math.Abs(baseline[0].Y - line.Y) > 1.5D) {
                if (TryCombineSameBaselineLines(baseline) is { } combined) rows.Add(combined);
                baseline.Clear();
            }
            baseline.Add(line);
        }
        if (baseline.Count > 0 && TryCombineSameBaselineLines(baseline) is { } last) rows.Add(last);
        var owned = new HashSet<PdfTextSpan>(tables.SelectMany(static table => table.SourceRuns));
        for (int index = 0; index + 2 < rows.Count; index++) {
            cancellationCheck?.Invoke();
            TextLayoutEngine.TextLine upper = rows[index], middle = rows[index + 1], lower = rows[index + 2];
            consumeWork?.Invoke(upper.Spans.Count + middle.Spans.Count + lower.Spans.Count);
            if (!HasEmphasizedText(upper)) continue;
            PositionedRow? first = TryCreatePositionedRow(upper), third = TryCreatePositionedRow(lower);
            if (first == null || third == null || !PositionedRowsAlign(first, third)) continue;
            double gap = upper.Y - middle.Y, nextGap = middle.Y - lower.Y;
            if (gap <= 0D || nextGap <= 0D || Math.Max(gap, nextGap) > Math.Min(gap, nextGap) * 1.35D ||
                Math.Max(gap, nextGap) > 48D) continue;
            double left = Math.Min(upper.XStart, lower.XStart), right = Math.Max(upper.XEnd, lower.XEnd);
            var contained = new List<PdfTextSpan>();
            bool crossing = false;
            foreach (PdfTextSpan span in middle.Spans) {
                double end = span.X + Math.Max(0D, span.Advance);
                if (span.X >= left - .25D && end <= right + .25D) contained.Add(span);
                else if (end > left && span.X < right) crossing = true;
            }
            if (crossing || contained.Count == 0 || contained.Count == middle.Spans.Count) continue;
            var bounded = new TextLayoutEngine.TextLine(middle.Y,
                contained.Min(static span => span.X), contained.Max(static span => span.X + Math.Max(0D, span.Advance)),
                ComposeCell(contained, middle.ReadingDirection), contained, middle.ReadingDirection);
            PositionedRow? second = TryCreatePositionedRow(bounded);
            if (second == null || !PositionedRowsAlign(first, second)) continue;
            var splits = new List<double>(first.Cells.Count - 1);
            for (int column = 1; column < first.Cells.Count; column++) {
                double previousEnd = Math.Max(first.Cells[column - 1].To,
                    Math.Max(second.Cells[column - 1].To, third.Cells[column - 1].To));
                double nextStart = Math.Min(first.Cells[column].From,
                    Math.Min(second.Cells[column].From, third.Cells[column].From));
                if (previousEnd >= nextStart) break;
                splits.Add((previousEnd + nextStart) / 2D);
            }
            if (splits.Count != first.Cells.Count - 1) continue;
            var source = new List<TextLayoutEngine.TextLine> { upper, bounded, lower };
            PdfStructuredTable? table = BuildTableFromLinesAndSplits(source, splits, "bounded-row-region");
            if (table == null || !HasValidatedRows(table, source)) continue;
            if (table.SourceRuns.Any(owned.Contains)) {
                PdfStructuredTable[] overlaps = tables.Where(existing =>
                    existing.SourceRuns.Any(table.SourceRuns.Contains)).ToArray();
                // A side annotation can also have been swallowed by an existing
                // three-row candidate. Replace only that complete region, never
                // remove a larger table or another independently owned region.
                if (overlaps.Length != 1 || overlaps[0].Rows.Count != 3 ||
                    overlaps[0].Columns.Count != table.Columns.Count ||
                    !table.SourceRuns.All(overlaps[0].SourceRuns.Contains)) continue;
                owned.ExceptWith(overlaps[0].SourceRuns);
                tables.Remove(overlaps[0]);
            }
            tables.Add(table);
            owned.UnionWith(table.SourceRuns);
        }
    }
}

namespace OfficeIMO.Pdf;

internal static partial class TableDetector {
    /// <summary>
    /// Recovers a compact header/body region when an unrelated side annotation changes
    /// the middle row's split count. A complete last-cell span group may extend
    /// past the other rows' paint when its start matches both last-column anchors.
    /// </summary>
    private static void RecoverTablesWithPeripheralRowText(List<PdfStructuredTable> tables,
        IReadOnlyList<TextLayoutEngine.TextLine> pageLines, double? pageHeight,
        Action<long>? consumeWork, Action? cancellationCheck) {
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
        var owned = new Dictionary<PdfTextSpan, int>();
        foreach (PdfStructuredTable existing in tables) ChangeOwnership(existing.SourceRuns, 1);
        for (int index = 0; index + 2 < rows.Count; index++) {
            cancellationCheck?.Invoke();
            TextLayoutEngine.TextLine upper = rows[index], middle = rows[index + 1], lower = rows[index + 2];
            TextLayoutEngine.TextLine? detachedNote = null;
            if (index + 3 < rows.Count && HasEmphasizedText(upper)) {
                TextLayoutEngine.TextLine finalRow = rows[index + 3];
                double regionLeft = Math.Min(upper.XStart, finalRow.XStart);
                double regionRight = Math.Max(upper.XEnd, finalRow.XEnd);
                consumeWork?.Invoke(middle.Spans.Count + lower.Spans.Count);
                bool middleOutside = IsOutside(middle), lowerOutside = IsOutside(lower);
                if (middleOutside != lowerOutside) {
                    TextLayoutEngine.TextLine note = middleOutside ? middle : lower;
                    TextLayoutEngine.TextLine body = middleOutside ? lower : middle;
                    double fontSize = body.Spans.Max(static span => span.FontSize);
                    if (Math.Abs(note.Y - body.Y) <= fontSize * .5D) {
                        var spans = body.Spans.Concat(note.Spans).OrderBy(static span => span.X).ToList();
                        middle = new TextLayoutEngine.TextLine(body.Y, spans.Min(static span => span.X),
                            spans.Max(static span => span.X + Math.Max(0D, span.Advance)),
                            ComposeCell(spans, body.ReadingDirection), spans, body.ReadingDirection);
                        lower = finalRow;
                        detachedNote = note;
                    }
                }

                bool IsOutside(TextLayoutEngine.TextLine line) => line.Spans.Count > 0 &&
                    (line.XEnd < regionLeft - .25D || line.XStart > regionRight + .25D);
            }
            consumeWork?.Invoke(upper.Spans.Count + middle.Spans.Count + lower.Spans.Count);
            if (!HasEmphasizedText(upper) || !CanRecoverTableLine(upper, pageHeight) ||
                !CanRecoverTableLine(lower, pageHeight)) continue;
            PositionedRow? first = TryCreatePositionedRow(upper), third = TryCreatePositionedRow(lower);
            if (first == null || third == null || !PositionedRowsAlign(first, third)) continue;
            double gap = upper.Y - middle.Y, nextGap = middle.Y - lower.Y;
            if (gap <= 0D || nextGap <= 0D || Math.Max(gap, nextGap) > Math.Min(gap, nextGap) * 1.35D ||
                Math.Max(gap, nextGap) > 48D) continue;
            double left = Math.Min(upper.XStart, lower.XStart), right = Math.Max(upper.XEnd, lower.XEnd);
            double firstLastStart = first.Cells[first.Cells.Count - 1].From;
            double thirdLastStart = third.Cells[third.Cells.Count - 1].From;
            var lastCellRuns = new HashSet<PdfTextSpan>();
            if (middle.XEnd > right + .25D) {
                consumeWork?.Invoke(middle.Spans.Count);
                PositionedRow? annotated = TryCreatePositionedRow(middle, useCompactGapThreshold: true);
                if (annotated == null) continue;
                foreach (PositionedCell cell in annotated.Cells) {
                    cancellationCheck?.Invoke();
                    consumeWork?.Invoke(1);
                    if (Math.Abs(cell.From - firstLastStart) > .25D ||
                        Math.Abs(cell.From - thirdLastStart) > .25D) continue;
                    consumeWork?.Invoke(cell.SourceRuns.Count);
                    foreach (PdfTextSpan span in cell.SourceRuns) lastCellRuns.Add(span);
                }
            }
            var contained = new List<PdfTextSpan>();
            bool crossing = false;
            foreach (PdfTextSpan span in middle.Spans) {
                double end = span.X + Math.Max(0D, span.Advance);
                if (span.X >= left - .25D && (end <= right + .25D || lastCellRuns.Contains(span))) {
                    if (!span.CanProjectCompleteText(pageHeight)) { crossing = true; break; }
                    contained.Add(span);
                }
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
            var candidateRuns = new HashSet<PdfTextSpan>();
            bool sharesOwnership = false;
            foreach (PdfTextSpan span in table.SourceRuns) {
                AccountWork();
                candidateRuns.Add(span);
                sharesOwnership |= owned.ContainsKey(span);
            }
            if (sharesOwnership) {
                var overlaps = new List<PdfStructuredTable>();
                foreach (PdfStructuredTable existing in tables) {
                    AccountWork();
                    foreach (PdfTextSpan span in existing.SourceRuns) {
                        AccountWork();
                        if (!candidateRuns.Contains(span)) continue;
                        overlaps.Add(existing);
                        break;
                    }
                    if (overlaps.Count > 1) break;
                }
                // A fourth row may consist solely of the nearby outside annotation.
                // Require exact source ownership before removing that row; larger or
                // separately owned table regions remain untouched.
                if (overlaps.Count != 1 || overlaps[0].Rows.Count < 2 ||
                    overlaps[0].Rows.Count > (detachedNote == null ? 3 : 4) ||
                    overlaps[0].Columns.Count != table.Columns.Count) continue;
                var previousRuns = new HashSet<PdfTextSpan>();
                foreach (PdfTextSpan span in overlaps[0].SourceRuns) {
                    AccountWork();
                    previousRuns.Add(span);
                }
                if (overlaps[0].Rows.Count == 4) {
                    var annotatedRuns = new HashSet<PdfTextSpan>(candidateRuns);
                    foreach (PdfTextSpan span in detachedNote!.Spans) {
                        AccountWork();
                        annotatedRuns.Add(span);
                    }
                    AccountWork(annotatedRuns.Count + previousRuns.Count);
                    if (!annotatedRuns.SetEquals(previousRuns)) continue;
                }
                var requiredRuns = overlaps[0].Rows.Count == 2 ? previousRuns : candidateRuns;
                var containingRuns = overlaps[0].Rows.Count == 2 ? candidateRuns : previousRuns;
                bool ownershipMatches = true;
                foreach (PdfTextSpan span in requiredRuns) {
                    AccountWork();
                    if (!containingRuns.Contains(span)) { ownershipMatches = false; break; }
                }
                if (!ownershipMatches) continue;
                ChangeOwnership(overlaps[0].SourceRuns, -1);
                AccountWork(tables.Count);
                tables.Remove(overlaps[0]);
            }
            ChangeOwnership(table.SourceRuns, 1);
            cancellationCheck?.Invoke();
            tables.Add(table);
        }

        void AccountWork(long units = 1) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(units);
        }

        void ChangeOwnership(IEnumerable<PdfTextSpan> spans, int change) {
            foreach (PdfTextSpan span in spans) {
                AccountWork();
                owned.TryGetValue(span, out int count);
                int next = count + change;
                if (next <= 0) owned.Remove(span);
                else owned[span] = next;
            }
        }
    }
}

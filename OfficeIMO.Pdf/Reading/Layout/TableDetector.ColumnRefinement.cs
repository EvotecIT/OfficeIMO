namespace OfficeIMO.Pdf;

internal static partial class TableDetector {
    // Refine an existing table from compact column evidence without changing
    // ordinary page line grouping or joining unrelated same-baseline regions.
    private static void RecoverCompactColumnRefinements(List<PdfStructuredTable> tables,
        IReadOnlyList<TextLayoutEngine.TextLine> pageLines, Action<long>? consumeWork, Action? cancellationCheck) {
        for (int index = 0; index < tables.Count; index++) {
            cancellationCheck?.Invoke();
            PdfStructuredTable coarse = tables[index];
            if (coarse.Rows.Count < 2 || coarse.Columns.Count < 2 || coarse.Kind == "leaders" ||
                coarse.SourceRuns.Count == 0 || coarse.SourceRuns.Count > MaximumPositionedRecoveryCells) continue;
            var options = new TextLayoutEngine.Options {
                ReadingDirection = coarse.SourceLines.Count > 0 ? coarse.SourceLines[0].ReadingDirection : PdfReadingDirection.Auto
            };
            List<TextLayoutEngine.TextLine> ownedLines = TextLayoutEngine.BuildLines(
                coarse.SourceRuns, options, consumeWork, cancellationCheck);
            if (ownedLines.Count < 2 || !HasEmphasizedText(ownedLines[0])) continue;
            PositionedRow? header = TryCreatePositionedRow(ownedLines[0], useCompactGapThreshold: true);
            if (header == null || header.Cells.Count <= coarse.Columns.Count) continue;
            double left = coarse.SourceRuns.Min(static span => span.X);
            double right = coarse.SourceRuns.Max(static span => span.X + Math.Max(0D, span.Advance));
            var source = new HashSet<PdfTextSpan>(coarse.SourceRuns);
            // A coarse band may also have stopped before an aligned final row.
            // Limit recovery to its horizontal region, then stop at the first
            // incompatible row, header, or vertical gap below its own header.
            foreach (TextLayoutEngine.TextLine line in pageLines) {
                cancellationCheck?.Invoke();
                if (line.Y > ownedLines[0].Y + 1.5D || line.Spans.Count == 0) continue;
                consumeWork?.Invoke(line.Spans.Count);
                if (!line.Spans.Any(span => span.X <= right + 1D && span.X + Math.Max(0D, span.Advance) >= left - 1D)) continue;
                // Occupied text width is not the table's physical cell boundary.
                // Keep the complete source fragment, including a longer final cell.
                foreach (PdfTextSpan span in line.Spans) {
                    source.Add(span);
                    if (source.Count > MaximumPositionedRecoveryCells) break;
                }
                if (source.Count > MaximumPositionedRecoveryCells) break;
            }
            if (source.Count > MaximumPositionedRecoveryCells) continue;
            List<TextLayoutEngine.TextLine> lines = TextLayoutEngine.BuildLines(source.ToArray(), options, consumeWork, cancellationCheck);
            var anchors = new List<(TextLayoutEngine.TextLine Line, PositionedRow Row)>();
            foreach (TextLayoutEngine.TextLine line in lines) {
                cancellationCheck?.Invoke();
                consumeWork?.Invoke(line.Spans.Count);
                PositionedRow? row = TryCreatePositionedRow(line, useCompactGapThreshold: true);
                if (row == null || row.Cells.Count != header.Cells.Count ||
                    anchors.Count > 0 && (HasEmphasizedText(line) || anchors[anchors.Count - 1].Row.Y - row.Y > 36D) ||
                    anchors.Count > 1 && !PositionedRowsAlign(anchors[1].Row, row)) break;
                anchors.Add((line, row));
            }
            if (anchors.Count < 2) continue;
            // Left-aligned labels and right-aligned values need not share the
            // same X anchor. The bounded builder checks their combined cells.
            var detailed = new List<PdfStructuredTable>();
            TryAddWrappedPositionedCellTable(detailed, anchors, anchors.Select(static item => item.Line).ToList(), consumeWork, cancellationCheck,
                allowSingleBodyRow: anchors.Count == 2, allowUnwrappedHeader: true);
            PdfTextSpan[] requiredRuns = anchors.SelectMany(static item => item.Line.Spans).Distinct().ToArray();
            PdfStructuredTable? refinement = detailed.FirstOrDefault(candidate =>
                IsMoreDetailedColumnPartition(coarse, candidate, requiredRuns, consumeWork, cancellationCheck));
            if (refinement != null) tables[index] = refinement;
        }
    }

    // A single numeric body row has no repeated body X anchors. Its bold
    // header and separable combined cell bounds provide the stronger evidence.
    private static bool RecoverStrongTwoRowTables(List<PdfStructuredTable> tables,
        IReadOnlyList<TextLayoutEngine.TextLine> pageLines, Action<long>? consumeWork, Action? cancellationCheck) {
        if (!TryCollectUnclaimedRecoveryRuns(tables, pageLines, consumeWork, cancellationCheck, out PdfTextSpan[] spans)) return false;
        bool recovered = false;
        List<TextLayoutEngine.TextLine> lines = TextLayoutEngine.BuildLines(spans,
            new TextLayoutEngine.Options(), consumeWork, cancellationCheck);
        for (int index = 0; index + 1 < lines.Count; index++) {
            cancellationCheck?.Invoke();
            TextLayoutEngine.TextLine headerLine = lines[index];
            TextLayoutEngine.TextLine bodyLine = lines[index + 1];
            if (!HasEmphasizedText(headerLine) || HasEmphasizedText(bodyLine) ||
                headerLine.Y - bodyLine.Y <= 0D || headerLine.Y - bodyLine.Y > 36D) continue;
            PositionedRow? header = TryCreatePositionedRow(headerLine, useCompactGapThreshold: true);
            PositionedRow? body = TryCreatePositionedRow(bodyLine, useCompactGapThreshold: true);
            if (header == null || body == null || header.Cells.Count != body.Cells.Count ||
                !LooksLikeHeaderRow(header.Cells.Select(static cell => cell.Text).ToArray()) ||
                !body.Cells.Any(static cell => IsTabularValue(cell.Text))) continue;
            var anchors = new List<(TextLayoutEngine.TextLine Line, PositionedRow Row)> {
                (headerLine, header), (bodyLine, body)
            };
            var candidates = new List<PdfStructuredTable>();
            TryAddWrappedPositionedCellTable(candidates, anchors, new List<TextLayoutEngine.TextLine> { headerLine, bodyLine },
                consumeWork, cancellationCheck, allowSingleBodyRow: true, allowUnwrappedHeader: true);
            if (candidates.Count != 1) continue;
            PdfStructuredTable candidate = candidates[0];
            var candidateRuns = new HashSet<PdfTextSpan>(candidate.SourceRuns);
            consumeWork?.Invoke(headerLine.Spans.Count + bodyLine.Spans.Count);
            if (!headerLine.Spans.All(candidateRuns.Contains) || !bodyLine.Spans.All(candidateRuns.Contains) ||
                !TryReleaseTrailingRecoveryHeader(tables, headerLine.Spans, candidate.Rows[0], consumeWork, cancellationCheck)) continue;
            tables.Add(candidate);
            recovered = true;
        }
        return recovered;
    }

    // A coarse band can consume the next table's bold header as its last row.
    // Transfer only that exact row after a complete numeric body is recovered.
    private static bool TryReleaseTrailingRecoveryHeader(List<PdfStructuredTable> tables,
        IReadOnlyList<PdfTextSpan> headerRuns, string[] header, Action<long>? consumeWork, Action? cancellationCheck) {
        var headerSet = new HashSet<PdfTextSpan>(headerRuns);
        var owners = new List<PdfStructuredTable>();
        foreach (PdfStructuredTable table in tables) {
            cancellationCheck?.Invoke();
            if (table.SourceRuns.Count > 0) consumeWork?.Invoke(table.SourceRuns.Count);
            if (!table.SourceRuns.Any(headerSet.Contains)) continue;
            if (table.Kind == "leaders" || table.Rows.Count < 3 || table.SourceLines.Count != table.Rows.Count) return false;
            TextLayoutEngine.TextLine lastLine = table.SourceLines[table.SourceLines.Count - 1];
            if (!headerSet.SetEquals(lastLine.Spans) ||
                !string.Equals(string.Join(" ", table.Rows[table.Rows.Count - 1]), string.Join(" ", header), StringComparison.Ordinal)) return false;
            owners.Add(table);
        }
        foreach (PdfStructuredTable table in owners) {
            table.Rows.RemoveAt(table.Rows.Count - 1);
            table.SourceLines = table.SourceLines.Take(table.SourceLines.Count - 1).ToArray();
            table.SourceRuns = table.SourceRuns.Where(span => !headerSet.Contains(span)).ToArray();
            table.YBottom = table.SourceLines[table.SourceLines.Count - 1].Y;
        }
        return true;
    }

    // Keep claimed body fragments out of new recovery regions. Emphasized
    // header fragments remain eligible for exact trailing-row reconciliation.
    private static bool TryCollectUnclaimedRecoveryRuns(List<PdfStructuredTable> tables,
        IReadOnlyList<TextLayoutEngine.TextLine> pageLines, Action<long>? consumeWork, Action? cancellationCheck,
        out PdfTextSpan[] spans) {
        spans = Array.Empty<PdfTextSpan>();
        var claimed = new HashSet<PdfTextSpan>();
        int inspected = 0;
        foreach (PdfStructuredTable table in tables) {
            cancellationCheck?.Invoke();
            if (table.SourceRuns.Count > MaximumPositionedRecoveryCells - inspected) return false;
            inspected += table.SourceRuns.Count;
            if (table.SourceRuns.Count > 0) consumeWork?.Invoke(table.SourceRuns.Count);
            foreach (PdfTextSpan span in table.SourceRuns) claimed.Add(span);
        }
        var unclaimed = new HashSet<PdfTextSpan>();
        inspected = 0;
        foreach (TextLayoutEngine.TextLine line in pageLines) {
            cancellationCheck?.Invoke();
            if (line.Spans.Count > MaximumPositionedRecoveryCells - inspected) return false;
            inspected += line.Spans.Count;
            if (line.Spans.Count == 0) continue;
            consumeWork?.Invoke(line.Spans.Count);
            if (line.Spans.Any(claimed.Contains) && !HasEmphasizedText(line)) continue;
            foreach (PdfTextSpan span in line.Spans) unclaimed.Add(span);
        }
        spans = unclaimed.ToArray();
        return spans.Length > 0;
    }

    private static bool IsMoreDetailedColumnPartition(PdfStructuredTable coarse,
        PdfStructuredTable detailed, PdfTextSpan[] requiredRuns, Action<long>? consumeWork, Action? cancellationCheck) {
        if (detailed.Columns.Count <= coarse.Columns.Count || detailed.Rows.Count < coarse.Rows.Count ||
            detailed.Kind != "wrapped-positioned-cells-bounded") return false;
        cancellationCheck?.Invoke();
        consumeWork?.Invoke(detailed.SourceRuns.Count + coarse.SourceRuns.Count + requiredRuns.Length);
        var runs = new HashSet<PdfTextSpan>(detailed.SourceRuns);
        if (!coarse.SourceRuns.All(runs.Contains) || !requiredRuns.All(runs.Contains)) return false;
        int matched = 0;
        for (int row = 0; row < detailed.Rows.Count && matched < coarse.Rows.Count; row++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(coarse.Rows[matched].Sum(static cell => (long)cell.Length) +
                detailed.Rows[row].Sum(static cell => (long)cell.Length));
            string left = string.Join(" ", coarse.Rows[matched].Where(static cell => !string.IsNullOrWhiteSpace(cell)));
            string right = string.Join(" ", detailed.Rows[row].Where(static cell => !string.IsNullOrWhiteSpace(cell)));
            if (string.Equals(left, right, StringComparison.Ordinal)) matched++;
        }
        return matched == coarse.Rows.Count;
    }
}

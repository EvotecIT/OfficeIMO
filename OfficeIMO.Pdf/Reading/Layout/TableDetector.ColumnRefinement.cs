namespace OfficeIMO.Pdf;

internal static partial class TableDetector {
    // Refine an existing table from compact column evidence without changing
    // ordinary page line grouping or joining unrelated same-baseline regions.
    private static void RecoverCompactColumnRefinements(List<PdfStructuredTable> tables,
        IReadOnlyList<TextLayoutEngine.TextLine> pageLines, Action<long>? consumeWork, Action? cancellationCheck) {
        for (int index = 0; index < tables.Count; index++) {
            cancellationCheck?.Invoke();
            PdfStructuredTable coarse = tables[index];
            if (coarse.Rows.Count < 3 || coarse.Columns.Count < 2 || coarse.Kind == "leaders" ||
                coarse.SourceRuns.Count == 0 || coarse.SourceRuns.Count > MaximumPositionedRecoveryCells) continue;
            var options = new TextLayoutEngine.Options {
                ReadingDirection = coarse.SourceLines.Count > 0 ? coarse.SourceLines[0].ReadingDirection : PdfReadingDirection.Auto
            };
            List<TextLayoutEngine.TextLine> ownedLines = TextLayoutEngine.BuildLines(
                coarse.SourceRuns, options, consumeWork, cancellationCheck);
            if (ownedLines.Count < 3 || !HasEmphasizedText(ownedLines[0])) continue;
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
                if (line.Y > ownedLines[0].Y + 1.5D) continue;
                foreach (PdfTextSpan span in line.Spans) {
                    consumeWork?.Invoke(1);
                    if (span.X >= left - 1D && span.X + Math.Max(0D, span.Advance) <= right + 1D) source.Add(span);
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
            if (anchors.Count < 3) continue;
            // Left-aligned labels and right-aligned values need not share the
            // same X anchor. The bounded builder checks their combined cells.
            var detailed = new List<PdfStructuredTable>();
            TryAddWrappedPositionedCellTable(detailed, anchors, anchors.Select(static item => item.Line).ToList(), consumeWork, cancellationCheck);
            PdfStructuredTable? refinement = detailed.FirstOrDefault(candidate =>
                IsMoreDetailedColumnPartition(coarse, candidate, consumeWork, cancellationCheck));
            if (refinement != null) tables[index] = refinement;
        }
    }

    // A single numeric body row has no repeated body X anchors. Its bold
    // header and separable combined cell bounds provide the stronger evidence.
    private static void RecoverStrongTwoRowTables(List<PdfStructuredTable> tables,
        IReadOnlyList<TextLayoutEngine.TextLine> pageLines, Action<long>? consumeWork, Action? cancellationCheck) {
        PdfTextSpan[] spans = pageLines.SelectMany(static line => line.Spans).Distinct()
            .Take(MaximumPositionedRecoveryCells + 1).ToArray();
        if (spans.Length == 0 || spans.Length > MaximumPositionedRecoveryCells) return;
        consumeWork?.Invoke(spans.Length);
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
            TryAddWrappedPositionedCellTable(tables, anchors, new List<TextLayoutEngine.TextLine> { headerLine, bodyLine },
                consumeWork, cancellationCheck, allowSingleBodyRow: true);
        }
    }

    private static bool IsMoreDetailedColumnPartition(PdfStructuredTable coarse,
        PdfStructuredTable detailed, Action<long>? consumeWork, Action? cancellationCheck) {
        if (detailed.Columns.Count <= coarse.Columns.Count || detailed.Rows.Count < coarse.Rows.Count ||
            detailed.Kind != "wrapped-positioned-cells-bounded") return false;
        cancellationCheck?.Invoke();
        consumeWork?.Invoke(detailed.SourceRuns.Count + coarse.SourceRuns.Count);
        var runs = new HashSet<PdfTextSpan>(detailed.SourceRuns);
        if (!coarse.SourceRuns.All(runs.Contains)) return false;
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

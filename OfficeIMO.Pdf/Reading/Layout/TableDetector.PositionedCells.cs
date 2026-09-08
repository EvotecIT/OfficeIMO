namespace OfficeIMO.Pdf;

internal static partial class TableDetector {
    internal static List<StructuredTable> DetectPositionedCellTables(
        IReadOnlyList<TextLayoutEngine.TextLine> lines,
        double? pageHeight = null,
        Action<long>? consumeWork = null,
        Action? cancellationCheck = null) {
        var result = new List<StructuredTable>();
        var group = new List<PositionedRow>();
        int inspectedLines = 0;
        int inspectedCells = 0;
        foreach (TextLayoutEngine.TextLine line in lines.OrderByDescending(static line => line.Y)) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            if (inspectedLines++ == MaximumPositionedRecoveryLines) break;
            if (!CanRecoverTableLine(line, pageHeight)) {
                AddPositionedGroup(result, group);
                group.Clear();
                continue;
            }
            PositionedRow? row = TryCreatePositionedRow(line);
            if (row == null || inspectedCells + row.Cells.Count > MaximumPositionedRecoveryCells) {
                AddPositionedGroup(result, group);
                group.Clear();
                if (row != null) break;
                continue;
            }

            inspectedCells += row.Cells.Count;
            if (group.Count > 0 &&
                (!PositionedRowsAlign(group[0], row) || HasLargeVerticalGap(group, row))) {
                AddPositionedGroup(result, group);
                group.Clear();
            }
            group.Add(row);
        }

        AddPositionedGroup(result, group);
        return result;
    }

    private static PositionedRow? TryCreatePositionedRow(TextLayoutEngine.TextLine line) {
        if (line.Spans.Count < 2) return null;
        var cells = new List<PositionedCell>();
        var builder = new System.Text.StringBuilder();
        var sourceRuns = new List<PdfTextSpan>();
        double from = 0D;
        double to = 0D;
        double lastSpanStart = 0D;
        for (int index = 0; index < line.Spans.Count; index++) {
            PdfTextSpan span = line.Spans[index];
            bool split = false;
            double gap = 0D;
            if (index > 0) {
                PdfTextSpan previous = line.Spans[index - 1];
                double previousEnd = previous.X + Math.Max(0D, previous.Advance);
                gap = span.X - previousEnd;
                split = gap > Math.Max(18D, Math.Max(previous.FontSize, span.FontSize) * 2D);
            }

            if (split) {
                cells.Add(new PositionedCell(from, to, lastSpanStart, ComposeCell(sourceRuns), sourceRuns.ToArray()));
                builder.Clear();
                sourceRuns.Clear();
            } else if (gap > 1D && builder.Length > 0 && builder[builder.Length - 1] != ' ') {
                builder.Append(' ');
            }

            if (builder.Length == 0) from = span.X;
            builder.Append(span.Text);
            sourceRuns.Add(span);
            to = span.X + Math.Max(0D, span.Advance);
            lastSpanStart = span.X;
            if (cells.Count == MaximumPositionedRecoveryColumns) return null;
        }

        if (builder.Length > 0) cells.Add(new PositionedCell(from, to, lastSpanStart, ComposeCell(sourceRuns), sourceRuns.ToArray()));
        return cells.Count is >= 2 and <= MaximumPositionedRecoveryColumns
            ? new PositionedRow(line.Y, cells)
            : null;
    }

    private static bool PositionedRowsAlign(PositionedRow expected, PositionedRow current) {
        if (expected.Cells.Count != current.Cells.Count) return false;
        for (int index = 0; index < expected.Cells.Count; index++) {
            PositionedCell expectedCell = expected.Cells[index];
            PositionedCell currentCell = current.Cells[index];
            bool leftAligned = Math.Abs(expectedCell.From - currentCell.From) <= 16D;
            bool centerAligned = Math.Abs(
                (expectedCell.From + expectedCell.To) / 2D -
                (currentCell.From + currentCell.To) / 2D) <= 16D;
            bool rightAligned = Math.Abs(expectedCell.To - currentCell.To) <= 16D;
            if (!leftAligned && !centerAligned && !rightAligned) return false;
        }
        return true;
    }

    private static bool HasLargeVerticalGap(List<PositionedRow> rows, PositionedRow current) {
        if (rows.Count < 2) return false;
        double gap = rows[rows.Count - 1].Y - current.Y;
        if (gap <= 36D) return false;

        var priorGaps = new List<double>(rows.Count - 1);
        for (int index = 1; index < rows.Count; index++) {
            double priorGap = rows[index - 1].Y - rows[index].Y;
            if (priorGap > 0D) priorGaps.Add(priorGap);
        }
        if (priorGaps.Count == 0) return gap > 48D;
        priorGaps.Sort();
        double median = priorGaps[priorGaps.Count / 2];
        return gap > Math.Max(36D, median * 2.5D);
    }

    private static void AddPositionedGroup(List<StructuredTable> result, List<PositionedRow> rows) {
        if (rows.Count < 3 || !LooksLikePositionedTable(rows)) return;
        if (TryPartitionPositionedRows(rows, out List<PositionedRow>? left, out List<PositionedRow>? right)) {
            AddPositionedGroup(result, left);
            AddPositionedGroup(result, right);
            return;
        }
        var table = new StructuredTable {
            YTop = rows[0].Y,
            YBottom = rows[rows.Count - 1].Y,
            Kind = "positioned-cells-bounded"
        };
        for (int columnIndex = 0; columnIndex < rows[0].Cells.Count; columnIndex++) {
            table.Columns.Add(new StructuredTableColumn {
                From = rows.Min(row => row.Cells[columnIndex].From),
                To = rows.Max(row => row.Cells[columnIndex].To)
            });
        }
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            table.Rows.Add(rows[rowIndex].Cells.Select(static cell => cell.Text).ToArray());
        }
        table.SourceRuns = rows
            .SelectMany(static row => row.Cells)
            .SelectMany(static cell => cell.SourceRuns)
            .Distinct()
            .ToArray();
        if (!HasNarrativeColumnEvidence(table)) result.Add(table);
    }

    private static bool TryPartitionPositionedRows(
        List<PositionedRow> rows,
        out List<PositionedRow> left,
        out List<PositionedRow> right) {
        left = new List<PositionedRow>();
        right = new List<PositionedRow>();
        int columnCount = rows[0].Cells.Count;
        if (columnCount < 4 || rows.Any(row => row.Cells.Count != columnCount)) return false;

        int bestSplit = -1;
        double bestRatio = 0D;
        for (int split = 2; split <= columnCount - 2; split++) {
            var boundaryGaps = new List<double>(rows.Count);
            var otherGaps = new List<double>(rows.Count * Math.Max(1, columnCount - 2));
            foreach (PositionedRow row in rows) {
                for (int index = 1; index < row.Cells.Count; index++) {
                    double gap = Math.Max(0D, row.Cells[index].From - row.Cells[index - 1].To);
                    if (index == split) boundaryGaps.Add(gap);
                    else otherGaps.Add(gap);
                }
            }

            double boundary = Median(boundaryGaps);
            double typical = Median(otherGaps);
            if (boundary < Math.Max(72D, typical * 2D)) continue;
            double ratio = boundary / Math.Max(1D, typical);
            if (ratio > bestRatio) {
                bestRatio = ratio;
                bestSplit = split;
            }
        }
        if (bestSplit < 0) return false;

        foreach (PositionedRow row in rows) {
            left.Add(new PositionedRow(row.Y, row.Cells.Take(bestSplit).ToList()));
            right.Add(new PositionedRow(row.Y, row.Cells.Skip(bestSplit).ToList()));
        }
        if (LooksLikePositionedTable(left) && LooksLikePositionedTable(right)) return true;
        left.Clear();
        right.Clear();
        return false;
    }

    private static double Median(List<double> values) {
        if (values.Count == 0) return 0D;
        values.Sort();
        int middle = values.Count / 2;
        return (values.Count & 1) == 0
            ? (values[middle - 1] + values[middle]) / 2D
            : values[middle];
    }

    private static bool LooksLikePositionedTable(List<PositionedRow> rows) {
        string[] header = rows[0].Cells.Select(static cell => cell.Text).ToArray();
        if (!LooksLikeHeaderRow(header)) return false;
        for (int rowIndex = 1; rowIndex < rows.Count; rowIndex++) {
            for (int columnIndex = 0; columnIndex < rows[rowIndex].Cells.Count; columnIndex++) {
                string value = rows[rowIndex].Cells[columnIndex].Text;
                if (HasManyDigits(value)) return true;
            }
        }
        return HasTextualGridEvidence(rows);
    }

    private static bool HasTextualGridEvidence(List<PositionedRow> rows) {
        // Header plus at least three compact, consistently aligned body rows is
        // language-neutral evidence for categorical tables without numeric cells.
        if (rows.Count < 4) return false;
        int columnCount = rows[0].Cells.Count;
        if (columnCount < 2 || rows.Any(row => row.Cells.Count != columnCount)) return false;

        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                PositionedCell cell = rows[rowIndex].Cells[columnIndex];
                string value = ContentStructureExtractor.NormalizeShattered(cell.Text).Trim();
                if (value.Length == 0 ||
                    cell.OccupiedWidthInFontSizes > MaximumCompactCellWidthInFontSizes) {
                    return false;
                }
            }
        }

        var gaps = new List<double>(rows.Count - 1);
        for (int rowIndex = 1; rowIndex < rows.Count; rowIndex++) {
            double gap = rows[rowIndex - 1].Y - rows[rowIndex].Y;
            if (gap <= 0D) return false;
            gaps.Add(gap);
        }
        double medianGap = Median(gaps);
        double tolerance = Math.Max(3D, medianGap * 0.35D);
        return gaps.All(gap => Math.Abs(gap - medianGap) <= tolerance);
    }

    private sealed class PositionedRow {
        internal PositionedRow(double y, List<PositionedCell> cells) {
            Y = y;
            Cells = cells;
        }
        internal double Y { get; }
        internal List<PositionedCell> Cells { get; }
    }

    private readonly struct PositionedCell {
        internal PositionedCell(
            double from,
            double to,
            double lastSpanStart,
            string text,
            IReadOnlyList<PdfTextSpan> sourceRuns) {
            From = from;
            To = to;
            LastSpanStart = lastSpanStart;
            Text = text;
            SourceRuns = sourceRuns;
        }
        internal double From { get; }
        internal double To { get; }
        internal double LastSpanStart { get; }
        internal string Text { get; }
        internal IReadOnlyList<PdfTextSpan> SourceRuns { get; }
        internal double OccupiedWidthInFontSizes {
            get {
                double fontSize = 0D;
                for (int index = 0; index < SourceRuns.Count; index++) {
                    fontSize = Math.Max(fontSize, SourceRuns[index].FontSize);
                }
                return fontSize > 0D
                    ? Math.Max(0D, To - From) / fontSize
                    : double.PositiveInfinity;
            }
        }
    }

}

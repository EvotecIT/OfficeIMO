namespace OfficeIMO.Pdf;

/// <summary>
/// Recovers table evidence from normalized OCR geometry inside the canonical table-detection stage.
/// It does not interpret natural-language cell values.
/// </summary>
internal static class PdfOcrTableEvidenceDetector {
    private const int MinimumAlignedRows = 3;
    private const double MinimumColumnGapPoints = 18D;
    private const double MinimumColumnTolerancePoints = 12D;
    private const double MaximumCellWidthInTextHeights = 24D;
    private const double MaximumAverageCellWidthInTextHeights = 10D;

    internal static IReadOnlyList<PdfUnderstandingTableCandidate> Detect(
        PdfUnderstandingPageContext context,
        IReadOnlyList<PdfUnderstandingLine> lines,
        int maximumTables) {
        List<VisualRow> rows = BuildRows(context, lines);
        var candidates = new List<(int RowIndex, IReadOnlyList<VisualCell> Cells)>();
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            context.ConsumeWork();
            IReadOnlyList<VisualCell> cells = SplitCells(context, rows[rowIndex]);
            if (cells.Count >= 2) candidates.Add((rowIndex, cells));
        }
        if (candidates.Count < MinimumAlignedRows) return Array.Empty<PdfUnderstandingTableCandidate>();

        var groups = new List<List<(int RowIndex, IReadOnlyList<VisualCell> Cells)>>();
        for (int candidateIndex = 0; candidateIndex < candidates.Count; candidateIndex++) {
            context.ConsumeWork();
            (int rowIndex, IReadOnlyList<VisualCell> cells) = candidates[candidateIndex];
            List<(int RowIndex, IReadOnlyList<VisualCell> Cells)>? group = groups.LastOrDefault();
            bool followsPrevious = group is not null && rowIndex == group[group.Count - 1].RowIndex + 1;
            bool aligned = group is not null && ColumnsAlign(group[0].Cells, cells);
            bool compactGap = group is not null && HasCompactGap(rows[group[group.Count - 1].RowIndex], rows[rowIndex]);
            if (!followsPrevious || !aligned || !compactGap) {
                group = new List<(int RowIndex, IReadOnlyList<VisualCell> Cells)>();
                groups.Add(group);
            }
            group!.Add((rowIndex, cells));
        }

        var result = new List<PdfUnderstandingTableCandidate>();
        for (int groupIndex = 0; groupIndex < groups.Count; groupIndex++) {
            context.ThrowIfCancellationRequested();
            List<(int RowIndex, IReadOnlyList<VisualCell> Cells)> group = groups[groupIndex];
            if (group.Count < MinimumAlignedRows || !HasConservativeTableEvidence(group, rows)) continue;
            if (result.Count >= maximumTables) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts, maximumTables, result.Count + 1L);
            }

            int columnCount = group[0].Cells.Count;
            var columnBounds = new (double From, double To)[columnCount];
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                context.ConsumeWork(group.Count);
                columnBounds[columnIndex] = (
                    group.Min(row => row.Cells[columnIndex].Left),
                    group.Max(row => row.Cells[columnIndex].Right));
            }
            IReadOnlyList<IReadOnlyList<string>> tableRows = group
                .Select(row => (IReadOnlyList<string>)row.Cells.Select(static cell => cell.Text).ToArray())
                .ToArray();
            VisualRow[] sourceRows = group.Select(row => rows[row.RowIndex]).ToArray();
            double top = sourceRows.Min(static row => row.Top);
            double bottom = sourceRows.Max(static row => row.Bottom);
            double left = columnBounds.Min(static column => column.From);
            double right = columnBounds.Max(static column => column.To);
            PdfUnderstandingLine[] sourceLines = sourceRows
                .SelectMany(static row => row.Lines)
                .Distinct()
                .ToArray();
            double confidence = PdfInference.Clamp(sourceLines.Average(static line => line.Confidence));
            result.Add(PdfUnderstandingTableCandidate.FromOcr(
                "ocr-aligned-geometry",
                top,
                bottom,
                new PdfLogicalVisualBounds(left, top, right, bottom),
                columnBounds,
                tableRows,
                confidence,
                new[] { new PdfInferenceEvidence(
                    "table.ocr-aligned-geometry",
                    "Normalized OCR geometry forms repeated columns with compact row rhythm.",
                    0.85D) },
                sourceLines));
        }
        return result.Count == 0 ? Array.Empty<PdfUnderstandingTableCandidate>() : result.AsReadOnly();
    }

    private static List<VisualRow> BuildRows(
        PdfUnderstandingPageContext context,
        IReadOnlyList<PdfUnderstandingLine> lines) {
        PdfUnderstandingLine[] candidates = lines
            .Where(static line => line.SourceKind == PdfLogicalContentSourceKind.Ocr && line.VisualBounds is not null)
            .OrderBy(static line => line.VisualBounds!.Top)
            .ThenBy(static line => line.VisualBounds!.Left)
            .ToArray();
        var rows = new List<VisualRow>();
        for (int lineIndex = 0; lineIndex < candidates.Length; lineIndex++) {
            context.ConsumeWork();
            PdfUnderstandingLine line = candidates[lineIndex];
            PdfLogicalVisualBounds bounds = line.VisualBounds!;
            double center = (bounds.Top + bounds.Bottom) / 2D;
            VisualRow? best = null;
            for (int rowIndex = rows.Count - 1; rowIndex >= 0; rowIndex--) {
                context.ConsumeWork();
                VisualRow row = rows[rowIndex];
                double maximumDistance = Math.Max(2D, Math.Min(row.Height, bounds.Height) * 0.6D);
                if (center - row.CenterY > maximumDistance) break;
                if (Math.Abs(row.CenterY - center) <= maximumDistance) {
                    best = row;
                    break;
                }
            }
            if (best is null) {
                best = new VisualRow();
                rows.Add(best);
            }
            best.Add(line);
        }
        return rows;
    }

    private static IReadOnlyList<VisualCell> SplitCells(
        PdfUnderstandingPageContext context,
        VisualRow row) {
        var sourceOrder = new Dictionary<PdfUnderstandingWord, int>();
        int sequence = 0;
        foreach (PdfUnderstandingLine line in row.Lines.OrderBy(static line => line.SourceSequence ?? int.MaxValue)) {
            for (int wordIndex = 0; wordIndex < line.Words.Count; wordIndex++) {
                context.ConsumeWork();
                if (!sourceOrder.ContainsKey(line.Words[wordIndex])) sourceOrder.Add(line.Words[wordIndex], sequence++);
            }
        }
        WordBox[] positioned = sourceOrder.Keys
            .Where(static word => word.VisualBounds is not null)
            .Select(word => new WordBox(word, word.VisualBounds!, sourceOrder[word]))
            .OrderBy(static word => word.Bounds.Left)
            .ThenBy(static word => word.SourceSequence)
            .ToArray();
        if (positioned.Length == 0) return Array.Empty<VisualCell>();

        var cells = new List<VisualCell>();
        var current = new List<WordBox> { positioned[0] };
        for (int wordIndex = 1; wordIndex < positioned.Length; wordIndex++) {
            context.ConsumeWork();
            WordBox previous = positioned[wordIndex - 1];
            WordBox word = positioned[wordIndex];
            double minimumGap = Math.Max(
                MinimumColumnGapPoints,
                Math.Min(previous.Bounds.Height, word.Bounds.Height) * 1.25D);
            if (word.Bounds.Left - previous.Bounds.Right >= minimumGap) {
                cells.Add(VisualCell.From(current));
                current.Clear();
            }
            current.Add(word);
        }
        cells.Add(VisualCell.From(current));
        return cells;
    }

    private static bool HasConservativeTableEvidence(
        IReadOnlyList<(int RowIndex, IReadOnlyList<VisualCell> Cells)> group,
        IReadOnlyList<VisualRow> rows) {
        long cellCount = 0L;
        double occupiedWidthInTextHeights = 0D;
        for (int rowIndex = 0; rowIndex < group.Count; rowIndex++) {
            IReadOnlyList<VisualCell> cells = group[rowIndex].Cells;
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                double compactness = cells[cellIndex].OccupiedWidthInTextHeights;
                if (compactness <= 0D || compactness > MaximumCellWidthInTextHeights) return false;
                cellCount++;
                occupiedWidthInTextHeights += compactness;
            }
        }
        if (cellCount == 0L ||
            occupiedWidthInTextHeights > cellCount * MaximumAverageCellWidthInTextHeights) return false;

        double[] steps = new double[group.Count - 1];
        for (int rowIndex = 1; rowIndex < group.Count; rowIndex++) {
            steps[rowIndex - 1] = rows[group[rowIndex].RowIndex].CenterY - rows[group[rowIndex - 1].RowIndex].CenterY;
            if (steps[rowIndex - 1] <= 0D) return false;
        }
        double medianStep = Median(steps);
        double medianHeight = Median(group.Select(row => rows[row.RowIndex].Height));
        return medianHeight > 0D &&
               medianStep <= Math.Max(24D, medianHeight * 3D) &&
               steps.Max() <= steps.Min() * 1.75D;
    }

    private static bool HasCompactGap(VisualRow previous, VisualRow current) {
        double step = current.CenterY - previous.CenterY;
        return step > 0D && step <= Math.Max(24D, Math.Max(previous.Height, current.Height) * 3D);
    }

    private static bool ColumnsAlign(IReadOnlyList<VisualCell> expected, IReadOnlyList<VisualCell> actual) {
        if (expected.Count != actual.Count) return false;
        for (int columnIndex = 0; columnIndex < expected.Count; columnIndex++) {
            double tolerance = Math.Max(
                MinimumColumnTolerancePoints,
                Math.Min(expected[columnIndex].Height, actual[columnIndex].Height) * 0.75D);
            if (Math.Abs(expected[columnIndex].Left - actual[columnIndex].Left) > tolerance) return false;
        }
        return true;
    }

    private static double Median(IEnumerable<double> values) {
        double[] ordered = values.OrderBy(static value => value).ToArray();
        if (ordered.Length == 0) return 1D;
        int middle = ordered.Length / 2;
        return ordered.Length % 2 == 0
            ? (ordered[middle - 1] + ordered[middle]) / 2D
            : ordered[middle];
    }

    private sealed class VisualRow {
        internal List<PdfUnderstandingLine> Lines { get; } = new();
        internal double Top { get; private set; }
        internal double Bottom { get; private set; }
        internal double CenterY => (Top + Bottom) / 2D;
        internal double Height => Bottom - Top;

        internal void Add(PdfUnderstandingLine line) {
            PdfLogicalVisualBounds bounds = line.VisualBounds!;
            if (Lines.Count == 0) {
                Top = bounds.Top;
                Bottom = bounds.Bottom;
            } else {
                Top = Math.Min(Top, bounds.Top);
                Bottom = Math.Max(Bottom, bounds.Bottom);
            }
            Lines.Add(line);
        }
    }

    private readonly struct WordBox {
        internal WordBox(PdfUnderstandingWord word, PdfLogicalVisualBounds bounds, int sourceSequence) {
            Word = word;
            Bounds = bounds;
            SourceSequence = sourceSequence;
        }

        internal PdfUnderstandingWord Word { get; }
        internal PdfLogicalVisualBounds Bounds { get; }
        internal int SourceSequence { get; }
    }

    private sealed class VisualCell {
        private VisualCell(double left, double right, double height, string text) {
            Left = left;
            Right = right;
            Height = height;
            Text = text;
        }

        internal double Left { get; }
        internal double Right { get; }
        internal double Height { get; }
        internal string Text { get; }
        internal double OccupiedWidthInTextHeights =>
            Height > 0D ? Math.Max(0D, Right - Left) / Height : double.PositiveInfinity;

        internal static VisualCell From(IReadOnlyList<WordBox> words) {
            WordBox[] sourceOrder = words.OrderBy(static word => word.SourceSequence).ToArray();
            return new VisualCell(
                words.Min(static word => word.Bounds.Left),
                words.Max(static word => word.Bounds.Right),
                words.Max(static word => word.Bounds.Bottom) - words.Min(static word => word.Bounds.Top),
                string.Join(" ", sourceOrder.Select(static word => word.Word.Text)));
        }
    }
}

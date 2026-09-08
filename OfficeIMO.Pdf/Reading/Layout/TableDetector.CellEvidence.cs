namespace OfficeIMO.Pdf;

internal static partial class TableDetector {
    private static bool LooksLikeCompactHeaderRow(StructuredTable table) {
        if (table.Rows.Count == 0 || !LooksLikeHeaderRow(table.Rows[0])) return false;
        double occupiedWidthInFontSizes = 0D;
        for (int columnIndex = 0; columnIndex < table.Rows[0].Length; columnIndex++) {
            string value = ContentStructureExtractor.NormalizeShattered(table.Rows[0][columnIndex]).Trim();
            if (ContentStructureExtractor.EndsWithSentenceTerminal(value) ||
                !TryGetOccupiedCellWidthInFontSizes(
                    table,
                    rowIndex: 0,
                    columnIndex,
                    out double compactness) ||
                compactness > MaximumCompactCellWidthInFontSizes) return false;
            occupiedWidthInFontSizes += compactness;
        }
        return occupiedWidthInFontSizes <=
               table.Rows[0].Length * MaximumAverageCompactCellWidthInFontSizes * 4D / 3D;
    }

    private static bool HasValidatedRows(StructuredTable table, IReadOnlyList<TextLayoutEngine.TextLine> sourceLines) {
        int columnCount = table.Columns.Count;
        if (columnCount < 2 || table.Rows.Count < 2) return false;

        int denseRows = 0;
        bool hasTabularValueEvidence = false;
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            string[] row = table.Rows[rowIndex];
            int populatedCells = 0;
            for (int columnIndex = 0; columnIndex < Math.Min(columnCount, row.Length); columnIndex++) {
                string value = row[columnIndex];
                if (!string.IsNullOrWhiteSpace(value)) populatedCells++;
                if (rowIndex > 0 && IsTabularValue(value)) hasTabularValueEvidence = true;
            }
            if (populatedCells >= 2 && populatedCells * 2 >= columnCount) denseRows++;
        }

        bool dense = denseRows >= 2 && denseRows * 2 >= table.Rows.Count;
        bool emphasizedHeader = HasEmphasizedHeader(sourceLines);
        if (HasNarrativeColumnEvidence(table)) return false;
        bool compactGrid = HasCompactCellGrid(table) && !HasPageColumnLikeGutters(sourceLines);
        return LooksLikeSparseFormGrid(table) ||
               (dense && (
                   hasTabularValueEvidence ||
                   emphasizedHeader ||
                   compactGrid ||
                   HasStableColumnAnchors(table, sourceLines)));
    }

    private static bool HasNarrativeColumnEvidence(StructuredTable table) {
        if (table.Columns.Count != 2 || table.SourceRuns.Count < 2) return false;
        int sentenceCells = table.Rows.Sum(row => row.Count(value =>
            ContentStructureExtractor.EndsWithSentenceTerminal(value.Trim())));
        if (sentenceCells < table.Rows.Count) return false;
        double[] baselines = table.SourceRuns.Select(static run => run.Y).Distinct().OrderBy(static value => value).ToArray();
        double fontSize = table.SourceRuns.Select(static span => span.FontSize).DefaultIfEmpty(1D).Max();
        if (table.SourceRuns.Where(run => Math.Abs(run.Y - baselines[baselines.Length - 1]) <= 2D)
            .Where(static run => !string.IsNullOrWhiteSpace(run.Text)).All(static run => run.IsBold)) return false;
        return baselines.Length >= 2 && baselines.Zip(baselines.Skip(1), static (lower, upper) => upper - lower)
            .All(gap => gap > fontSize * 2.2D);
    }

    private static bool LooksLikeSparseFormGrid(StructuredTable table) {
        if (table.Rows.Count < 3 ||
            !LooksLikeCompactHeaderRow(table)) return false;

        var sparseRowsByColumn = new int[table.Columns.Count];
        for (int rowIndex = 1; rowIndex < table.Rows.Count; rowIndex++) {
            string[] row = table.Rows[rowIndex];
            int populatedCells = 0;
            int populatedColumn = -1;
            for (int columnIndex = 0; columnIndex < Math.Min(row.Length, table.Columns.Count); columnIndex++) {
                if (string.IsNullOrWhiteSpace(row[columnIndex])) continue;
                populatedCells++;
                populatedColumn = columnIndex;
            }
            if (populatedCells != 1 || populatedColumn < 0) continue;

            string label = ContentStructureExtractor.NormalizeShattered(row[populatedColumn]).Trim();
            if (label.Length > 0 &&
                TryGetOccupiedCellWidthInFontSizes(
                    table,
                    rowIndex,
                    populatedColumn,
                    out double compactness) &&
                compactness <= MaximumCompactCellWidthInFontSizes * 1.25D &&
                !ContentStructureExtractor.EndsWithSentenceTerminal(label)) {
                sparseRowsByColumn[populatedColumn]++;
            }
        }
        return sparseRowsByColumn.Any(static count => count >= 2);
    }

    private static bool HasPageColumnLikeGutters(IReadOnlyList<TextLayoutEngine.TextLine> sourceLines) {
        if (sourceLines.Count < 3) return false;

        int separatedLines = 0;
        int inspectedLines = 0;
        for (int lineIndex = 0; lineIndex < sourceLines.Count; lineIndex++) {
            PdfTextSpan[] spans = sourceLines[lineIndex].Spans
                .Where(static span => !string.IsNullOrWhiteSpace(span.Text))
                .OrderBy(static span => span.X)
                .ToArray();
            if (spans.Length < 2) continue;

            inspectedLines++;
            double largestGap = 0D;
            var occupiedWidths = new List<double>(spans.Length);
            for (int spanIndex = 0; spanIndex < spans.Length; spanIndex++) {
                occupiedWidths.Add(Math.Max(1D, spans[spanIndex].Advance));
                if (spanIndex > 0) {
                    double previousRight = spans[spanIndex - 1].X + Math.Max(0D, spans[spanIndex - 1].Advance);
                    largestGap = Math.Max(largestGap, spans[spanIndex].X - previousRight);
                }
            }
            occupiedWidths.Sort();
            double medianOccupiedWidth = occupiedWidths[occupiedWidths.Count / 2];
            if (largestGap > Math.Max(72D, medianOccupiedWidth)) separatedLines++;
        }

        return inspectedLines >= 3 && separatedLines * 4 >= inspectedLines * 3;
    }

    private static double GetOccupiedWidthInFontSizes(TextLayoutEngine.TextLine line) {
        double left = double.MaxValue;
        double right = double.MinValue;
        double fontSize = 0D;
        for (int spanIndex = 0; spanIndex < line.Spans.Count; spanIndex++) {
            PdfTextSpan span = line.Spans[spanIndex];
            if (string.IsNullOrWhiteSpace(span.Text)) continue;
            double spanEnd = span.X + span.Advance;
            left = Math.Min(left, Math.Min(span.X, spanEnd));
            right = Math.Max(right, Math.Max(span.X, spanEnd));
            fontSize = Math.Max(fontSize, span.FontSize);
        }
        return fontSize > 0D && right >= left
            ? (right - left) / fontSize
            : double.PositiveInfinity;
    }

    private static bool TryGetOccupiedCellWidthInFontSizes(
        StructuredTable table,
        int rowIndex,
        int columnIndex,
        out double compactness) {
        compactness = double.PositiveInfinity;
        if (table.SourceLines.Count != table.Rows.Count ||
            rowIndex < 0 || rowIndex >= table.SourceLines.Count ||
            columnIndex < 0 || columnIndex >= table.Columns.Count) return false;

        StructuredTableColumn column = table.Columns[columnIndex];
        double columnLeft = Math.Min(column.From, column.To) - 0.5D;
        double columnRight = Math.Max(column.From, column.To) + 0.5D;
        TextLayoutEngine.TextLine line = table.SourceLines[rowIndex];
        double left = double.MaxValue;
        double right = double.MinValue;
        double fontSize = 0D;
        for (int spanIndex = 0; spanIndex < line.Spans.Count; spanIndex++) {
            PdfTextSpan span = line.Spans[spanIndex];
            if (string.IsNullOrWhiteSpace(span.Text) ||
                span.X < columnLeft || span.X > columnRight) continue;
            double spanEnd = span.X + span.Advance;
            left = Math.Min(left, Math.Min(span.X, spanEnd));
            right = Math.Max(right, Math.Max(span.X, spanEnd));
            fontSize = Math.Max(fontSize, span.FontSize);
        }
        if (fontSize <= 0D || right < left) return false;
        compactness = (right - left) / fontSize;
        return true;
    }

    private static bool HasStrongTwoRowEvidence(
        StructuredTable table,
        List<TextLayoutEngine.TextLine> sourceLines) {
        if (table.Rows.Count != 2 || sourceLines.Count != 2) return false;
        return HasStrongHeaderAndBodyEvidence(table, sourceLines);
    }

    private static bool HasStrongHeaderAndBodyEvidence(
        StructuredTable table,
        List<TextLayoutEngine.TextLine> sourceLines) {
        if (table.Rows.Count < 2 || sourceLines.Count < 2) return false;
        if (!LooksLikeHeaderRow(table.Rows[0])) return false;
        return table.Rows.Skip(1).Any(static row => row.Any(IsTabularValue)) ||
               (HasStableColumnAnchors(table, sourceLines) && HasEmphasizedHeader(sourceLines));
    }

    private static bool HasCompactCellGrid(StructuredTable table) {
        if (table.Rows.Count < 2 || !LooksLikeHeaderRow(table.Rows[0])) return false;
        int populatedCells = 0;
        double occupiedWidthInFontSizes = 0D;
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            string[] row = table.Rows[rowIndex];
            if (row.Length < table.Columns.Count) return false;
            for (int columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++) {
                string value = row[columnIndex].Trim();
                if (value.Length == 0) return false;
                if (!TryGetOccupiedCellWidthInFontSizes(
                        table,
                        rowIndex,
                        columnIndex,
                        out double compactness) ||
                    compactness > MaximumCompactCellWidthInFontSizes) return false;
                populatedCells++;
                occupiedWidthInFontSizes += compactness;
            }
        }
        return populatedCells > 0 &&
               occupiedWidthInFontSizes <= populatedCells * MaximumAverageCompactCellWidthInFontSizes;
    }

    private static bool HasStableColumnAnchors(
        StructuredTable table,
        IReadOnlyList<TextLayoutEngine.TextLine> sourceLines) {
        if (table.Rows.Count < 2 || !LooksLikeHeaderRow(table.Rows[0])) return false;
        for (int columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++) {
            StructuredTableColumn column = table.Columns[columnIndex];
            double left = Math.Min(column.From, column.To) - 0.5D;
            double right = Math.Max(column.From, column.To) + 0.5D;
            var leftAnchors = new List<double>();
            var centerAnchors = new List<double>();
            var rightAnchors = new List<double>();
            var fontSizes = new List<double>();
            for (int lineIndex = 0; lineIndex < sourceLines.Count; lineIndex++) {
                PdfTextSpan[] cellSpans = sourceLines[lineIndex].Spans
                    .Where(span => span.X >= left && span.X <= right && !string.IsNullOrWhiteSpace(span.Text))
                    .ToArray();
                if (cellSpans.Length == 0) continue;
                double cellLeft = cellSpans.Min(static span => span.X);
                double cellRight = cellSpans.Max(static span => span.X + Math.Max(0D, span.Advance));
                leftAnchors.Add(cellLeft);
                centerAnchors.Add((cellLeft + cellRight) / 2D);
                rightAnchors.Add(cellRight);
                fontSizes.Add(cellSpans.Max(static span => span.FontSize));
            }
            if (leftAnchors.Count < 2) return false;
            double tolerance = Math.Max(8D, Median(fontSizes));
            if (!HasStableAnchor(leftAnchors, tolerance) &&
                !HasStableAnchor(centerAnchors, tolerance) &&
                !HasStableAnchor(rightAnchors, tolerance)) return false;
        }
        return true;
    }

    private static bool HasStableAnchor(List<double> anchors, double tolerance) =>
        anchors.Max() - anchors.Min() <= tolerance;

    private static bool HasEmphasizedHeader(IReadOnlyList<TextLayoutEngine.TextLine> sourceLines) {
        if (sourceLines.Count < 2) return false;
        double headerY = sourceLines.Max(static line => line.Y);
        PdfTextSpan[] headerSpans = sourceLines
            .Where(line => Math.Abs(line.Y - headerY) <= 2D)
            .SelectMany(static line => line.Spans)
            .Where(static span => !string.IsNullOrWhiteSpace(span.Text))
            .ToArray();
        return headerSpans.Length >= 2 && headerSpans.All(static span => span.IsBold);
    }

    internal static bool HasDistinctEmphasizedHeader(IReadOnlyList<PdfUnderstandingLine> sourceLines) {
        if (sourceLines.Count < 2) return false;
        double headerY = sourceLines.Max(static line => line.BaselineY);
        PdfTextSpan[] headerSpans = sourceLines
            .Where(line => Math.Abs(line.BaselineY - headerY) <= 2D)
            .SelectMany(static line => line.Words)
            .SelectMany(static word => word.SourceRuns)
            .Where(static span => !string.IsNullOrWhiteSpace(span.Text))
            .Distinct()
            .ToArray();
        PdfTextSpan[] bodySpans = sourceLines
            .Where(line => line.BaselineY < headerY - 2D)
            .SelectMany(static line => line.Words)
            .SelectMany(static word => word.SourceRuns)
            .Where(static span => !string.IsNullOrWhiteSpace(span.Text))
            .Distinct()
            .ToArray();
        return headerSpans.Length >= 2 &&
               bodySpans.Length > 0 &&
               headerSpans.All(static span => span.IsBold) &&
               bodySpans.Any(static span => !span.IsBold);
    }

    private static bool IsTabularValue(string value) => HasManyDigits(value);

}

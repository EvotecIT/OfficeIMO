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
        AddWrappedPositionedCellTables(result, lines, pageHeight, consumeWork, cancellationCheck);
        return result;
    }

    private static void AddWrappedPositionedCellTables(
        List<StructuredTable> result,
        IReadOnlyList<TextLayoutEngine.TextLine> lines,
        double? pageHeight,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        var boundedLines = new List<TextLayoutEngine.TextLine>(Math.Min(lines.Count, MaximumPositionedRecoveryLines));
        var anchors = new List<(TextLayoutEngine.TextLine Line, PositionedRow Row)>();
        int inspectedCells = 0;
        for (int index = 0; index < lines.Count && boundedLines.Count < MaximumPositionedRecoveryLines; index++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            TextLayoutEngine.TextLine line = lines[index];
            boundedLines.Add(line);
            if (!CanRecoverTableLine(line, pageHeight)) continue;
            PositionedRow? row = TryCreatePositionedRow(line, useCompactGapThreshold: true);
            if (row is not { Cells.Count: >= 3 }) continue;
            if (inspectedCells + row.Cells.Count > MaximumPositionedRecoveryCells) break;
            inspectedCells += row.Cells.Count;
            anchors.Add((line, row));
        }
        anchors.Sort(static (left, right) => right.Row.Y.CompareTo(left.Row.Y));
        if (anchors.Count < 2) return;

        var group = new List<(TextLayoutEngine.TextLine Line, PositionedRow Row)>();
        for (int index = 0; index < anchors.Count; index++) {
            cancellationCheck?.Invoke();
            (TextLayoutEngine.TextLine Line, PositionedRow Row) candidate = anchors[index];
            if (group.Count > 0 &&
                (!PositionedRowsAlign(group[0].Row, candidate.Row) ||
                 group[group.Count - 1].Row.Y - candidate.Row.Y > 96D)) {
                TryAddWrappedPositionedCellTable(result, group, boundedLines, consumeWork, cancellationCheck);
                group.Clear();
            }
            group.Add(candidate);
        }
        TryAddWrappedPositionedCellTable(result, group, boundedLines, consumeWork, cancellationCheck);
    }

    private static void TryAddWrappedPositionedCellTable(
        List<StructuredTable> result,
        List<(TextLayoutEngine.TextLine Line, PositionedRow Row)> anchors,
        IReadOnlyList<TextLayoutEngine.TextLine> lines,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        cancellationCheck?.Invoke();
        if (anchors.Count < 2 ||
            anchors.Any(candidate => candidate.Row.Cells.Count != anchors[0].Row.Cells.Count)) return;

        int columnCount = anchors[0].Row.Cells.Count;
        if ((long)(anchors.Count + 1) * columnCount > MaximumPositionedRecoveryCells) return;
        consumeWork?.Invoke((long)anchors.Count * columnCount);
        double[] anchorGaps = anchors
            .Zip(anchors.Skip(1), static (upper, lower) => upper.Row.Y - lower.Row.Y)
            .Where(static gap => gap > 0D)
            .ToArray();
        if (anchorGaps.Length == 0) return;
        Array.Sort(anchorGaps);
        double rowPitch = anchorGaps[anchorGaps.Length / 2];
        if (rowPitch is < 12D or > 96D) return;

        double[] splits = new double[columnCount - 1];
        var occupiedFrom = new double[columnCount];
        var occupiedTo = new double[columnCount];
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            occupiedFrom[columnIndex] = anchors.Min(candidate => candidate.Row.Cells[columnIndex].From);
            occupiedTo[columnIndex] = anchors.Max(candidate => candidate.Row.Cells[columnIndex].To);
            if (columnIndex > 0) {
                double previousTo = occupiedTo[columnIndex - 1];
                double currentFrom = occupiedFrom[columnIndex];
                splits[columnIndex - 1] = (previousTo + currentFrom) / 2D;
            }
        }

        double left = occupiedFrom[0];
        double right = occupiedTo[columnCount - 1];
        var columns = new List<StructuredTableColumn>(columnCount);
        double columnLeft = left;
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            double columnRight = columnIndex < splits.Length ? splits[columnIndex] : right;
            columns.Add(new StructuredTableColumn { From = columnLeft, To = columnRight });
            columnLeft = columnRight;
        }
        string[] firstAnchorCells = anchors[0].Row.Cells
            .Select(static cell => ContentStructureExtractor.NormalizeShattered(cell.Text).Trim())
            .ToArray();
        bool firstAnchorIsHeader = anchors.Count >= 3 &&
            LooksLikeHeaderRow(firstAnchorCells) &&
            (HasEmphasizedText(anchors[0].Line) || !firstAnchorCells.Any(IsTabularValue));
        double headerBottom;
        List<TextLayoutEngine.TextLine> headerLines;
        int firstBodyAnchorIndex;
        if (firstAnchorIsHeader) {
            headerBottom = (anchors[0].Row.Y + anchors[1].Row.Y) / 2D;
            headerLines = new List<TextLayoutEngine.TextLine> { anchors[0].Line };
            firstBodyAnchorIndex = 1;
        } else {
            headerBottom = anchors[0].Row.Y + rowPitch / 2D;
            double headerTop = anchors[0].Row.Y + rowPitch;
            headerLines = SelectWrappedCellLines(
                lines,
                headerBottom,
                headerTop,
                left,
                right,
                consumeWork,
                cancellationCheck);
            if (headerLines.Count == 0) return;
            firstBodyAnchorIndex = 0;
        }
        (string[] Header, TextLayoutEngine.TextLine HeaderSource) = MergeWrappedCellLines(
            headerLines,
            splits,
            firstAnchorIsHeader ? anchors[0].Row.Y : anchors[0].Row.Y + rowPitch);
        if (!LooksLikeHeaderRow(Header)) return;

        var table = new StructuredTable {
            YTop = headerLines.Max(static line => line.Y),
            YBottom = anchors[anchors.Count - 1].Row.Y - rowPitch / 2D,
            Kind = "wrapped-positioned-cells-bounded"
        };
        table.Columns.AddRange(columns);
        table.Rows.Add(Header);
        var sourceLines = new List<TextLayoutEngine.TextLine> { HeaderSource };
        var sourceRuns = new List<PdfTextSpan>(HeaderSource.Spans);

        for (int anchorIndex = firstBodyAnchorIndex; anchorIndex < anchors.Count; anchorIndex++) {
            cancellationCheck?.Invoke();
            double upper = anchorIndex == 0
                ? headerBottom
                : (anchors[anchorIndex - 1].Row.Y + anchors[anchorIndex].Row.Y) / 2D;
            double lower = anchorIndex + 1 < anchors.Count
                ? (anchors[anchorIndex].Row.Y + anchors[anchorIndex + 1].Row.Y) / 2D
                : anchors[anchorIndex].Row.Y - rowPitch / 2D;
            List<TextLayoutEngine.TextLine> bodyLines = SelectWrappedCellLines(
                lines,
                lower,
                upper,
                left,
                right,
                consumeWork,
                cancellationCheck);
            if (bodyLines.Count == 0) return;
            (string[] Row, TextLayoutEngine.TextLine Source) = MergeWrappedCellLines(
                bodyLines,
                splits,
                anchors[anchorIndex].Row.Y);
            if (Row.Count(static cell => !string.IsNullOrWhiteSpace(cell)) != columnCount ||
                !Row.Any(IsTabularValue)) return;
            table.Rows.Add(Row);
            sourceLines.Add(Source);
            sourceRuns.AddRange(Source.Spans);
        }

        AppendTrailingPartialNumericRows(
            table,
            sourceLines,
            sourceRuns,
            lines,
            splits,
            left,
            right,
            anchors[anchors.Count - 1].Row.Y,
            rowPitch,
            consumeWork,
            cancellationCheck);

        table.SourceLines = sourceLines;
        table.SourceRuns = sourceRuns.Distinct().ToArray();
        if (!HasValidatedRows(table, sourceLines) ||
            result.Any(existing => IsSubsumedByPositionedTable(table, existing))) return;
        result.RemoveAll(existing => IsSubsumedByPositionedTable(existing, table));
        result.Add(table);
    }

    private static void AppendTrailingPartialNumericRows(
        StructuredTable table,
        List<TextLayoutEngine.TextLine> sourceLines,
        List<PdfTextSpan> sourceRuns,
        IReadOnlyList<TextLayoutEngine.TextLine> lines,
        double[] splits,
        double left,
        double right,
        double lastAnchorY,
        double rowPitch,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        int columnCount = splits.Length + 1;
        double currentAnchorY = lastAnchorY;
        while (table.Rows.Count < MaximumPositionedRecoveryLines &&
               (long)(table.Rows.Count + 1) * columnCount <= MaximumPositionedRecoveryCells &&
               TryFindTrailingNumericAnchor(
                   lines,
                   splits,
                   columnCount,
                   currentAnchorY,
                   rowPitch,
                   consumeWork,
                   cancellationCheck,
                   out TextLayoutEngine.TextLine numericAnchor)) {
            double upper = (currentAnchorY + numericAnchor.Y) / 2D;
            double lower = numericAnchor.Y - rowPitch / 2D;
            List<TextLayoutEngine.TextLine> bodyLines = SelectWrappedCellLines(
                lines,
                lower,
                upper,
                left,
                right,
                consumeWork,
                cancellationCheck);
            if (bodyLines.Count == 0) return;
            (string[] Row, TextLayoutEngine.TextLine Source) = MergeWrappedCellLines(
                bodyLines,
                splits,
                numericAnchor.Y);
            if (Row.Count(static cell => !string.IsNullOrWhiteSpace(cell)) != columnCount ||
                !Row.Any(IsTabularValue)) return;

            table.Rows.Add(Row);
            table.YBottom = lower;
            sourceLines.Add(Source);
            sourceRuns.AddRange(Source.Spans);
            currentAnchorY = numericAnchor.Y;
        }
    }

    private static bool TryFindTrailingNumericAnchor(
        IReadOnlyList<TextLayoutEngine.TextLine> lines,
        double[] splits,
        int columnCount,
        double currentAnchorY,
        double rowPitch,
        Action<long>? consumeWork,
        Action? cancellationCheck,
        out TextLayoutEngine.TextLine anchor) {
        anchor = null!;
        double expectedY = currentAnchorY - rowPitch;
        double tolerance = Math.Max(6D, rowPitch * 0.4D);
        double bestDistance = double.MaxValue;
        List<double> splitPositions = splits.ToList();
        for (int index = 0; index < lines.Count; index++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            TextLayoutEngine.TextLine line = lines[index];
            double distance = Math.Abs(line.Y - expectedY);
            if (distance > tolerance || distance >= bestDistance) continue;
            string[] cells = SplitBySplits(line, splitPositions);
            int firstPopulated = Array.FindIndex(cells, static cell => !string.IsNullOrWhiteSpace(cell));
            if (firstPopulated != columnCount - 3 ||
                PdfUnicodeScalarAnalysis.CountDecimalDigits(cells[firstPopulated]) == 0 ||
                cells.Skip(firstPopulated + 1).Any(static cell => string.IsNullOrWhiteSpace(cell) || !IsTabularValue(cell))) {
                continue;
            }
            anchor = line;
            bestDistance = distance;
        }
        return anchor != null;
    }

    private static List<TextLayoutEngine.TextLine> SelectWrappedCellLines(
        IReadOnlyList<TextLayoutEngine.TextLine> lines,
        double lowerY,
        double upperY,
        double left,
        double right,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        var selected = new List<TextLayoutEngine.TextLine>();
        for (int index = 0; index < lines.Count; index++) {
            cancellationCheck?.Invoke();
            consumeWork?.Invoke(1);
            TextLayoutEngine.TextLine line = lines[index];
            if (line.Y <= lowerY || line.Y > upperY) continue;
            List<PdfTextSpan> inBounds = line.Spans
                .Where(span => {
                    double spanEnd = span.X + Math.Max(0D, span.Advance);
                    return !string.IsNullOrWhiteSpace(span.Text) &&
                        spanEnd >= left - 4D &&
                        span.X <= right + 4D;
                })
                .ToList();
            if (inBounds.Count == 0) continue;
            selected.Add(inBounds.Count == line.Spans.Count
                ? line
                : new TextLayoutEngine.TextLine(
                    line.Y,
                    inBounds.Min(static span => span.X),
                    inBounds.Max(static span => span.X + Math.Max(0D, span.Advance)),
                    ComposeCell(inBounds, line.ReadingDirection),
                    inBounds,
                    line.ReadingDirection));
        }
        selected.Sort(static (upper, lower) => lower.Y.CompareTo(upper.Y));
        return selected;
    }

    private static (string[] Cells, TextLayoutEngine.TextLine SourceLine) MergeWrappedCellLines(
        List<TextLayoutEngine.TextLine> lines,
        double[] splits,
        double baselineY) {
        var builders = new System.Text.StringBuilder[splits.Length + 1];
        for (int columnIndex = 0; columnIndex < builders.Length; columnIndex++) {
            builders[columnIndex] = new System.Text.StringBuilder();
        }
        List<double> splitPositions = splits.ToList();
        foreach (TextLayoutEngine.TextLine line in lines) {
            string[] lineCells = SplitBySplits(line, splitPositions);
            for (int columnIndex = 0; columnIndex < lineCells.Length; columnIndex++) {
                string value = lineCells[columnIndex];
                if (value.Length == 0) continue;
                if (builders[columnIndex].Length > 0) builders[columnIndex].Append(' ');
                builders[columnIndex].Append(value);
            }
        }

        string[] cells = builders
            .Select(static builder => builder.ToString())
            .ToArray();
        List<PdfTextSpan> sourceRuns = lines.SelectMany(static line => line.Spans).ToList();
        return (cells, new TextLayoutEngine.TextLine(
            baselineY,
            sourceRuns.Min(static span => span.X),
            sourceRuns.Max(static span => span.X + Math.Max(0D, span.Advance)),
            string.Join(" ", cells.Where(static cell => cell.Length > 0)),
            sourceRuns));
    }

    private static PositionedRow? TryCreatePositionedRow(
        TextLayoutEngine.TextLine line,
        bool useCompactGapThreshold = false) {
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
                split = gap > (useCompactGapThreshold
                    ? GetCompactCellGapThreshold(previous, span)
                    : GetCellGapThreshold(previous, span));
            }

            if (split) {
                cells.Add(new PositionedCell(from, to, lastSpanStart, ComposeCell(sourceRuns, line.ReadingDirection), sourceRuns.ToArray()));
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

        if (builder.Length > 0) cells.Add(new PositionedCell(from, to, lastSpanStart, ComposeCell(sourceRuns, line.ReadingDirection), sourceRuns.ToArray()));
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

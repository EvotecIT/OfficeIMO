using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Pdf;

/// <summary>
/// Very simple table detector that splits a line into cells when there are large X gaps
/// between adjacent spans. Intended as a first cut for diagnostics and quick CSV-like rows.
/// Heuristics:
/// - Uses per-line span X coordinates and advances to compute inter-span gaps
/// - Splits when gap exceeds max(2*em, 18pt)
/// - Emits a row when at least two cells are produced and one cell is numeric-ish
/// </summary>
internal static class TableDetector {
    private const int MaximumPositionedRecoveryLines = 4096;
    private const int MaximumPositionedRecoveryColumns = 64;
    private const int MaximumPositionedRecoveryCells = 65536;
    public static List<string[]> Detect(List<TextLayoutEngine.TextLine> lines, double? pageHeight = null) {
        var rows = new List<string[]>();
        foreach (var match in DetectLineRows(lines, pageHeight)) {
            rows.Add(match.Cells);
        }
        return rows;
    }

    public static List<(TextLayoutEngine.TextLine Line, string[] Cells)> DetectLineRows(
        List<TextLayoutEngine.TextLine> lines,
        double? pageHeight = null) {
        var rows = new List<(TextLayoutEngine.TextLine Line, string[] Cells)>();
        foreach (var ln in lines) {
            if (!CanRecoverTableLine(ln, pageHeight) || ln.Spans.Count < 2) continue;
            var cells = SplitByGaps(ln);
            if (cells.Length >= 2 && LooksTabular(cells)) rows.Add((ln, cells));
        }
        return rows;
    }

    public static List<StructuredTable> DetectTablesFromBands(
        List<List<TextLayoutEngine.TextLine>> bands,
        double? pageHeight = null) {
        bands = bands
            .Select(band => band.Where(line => CanRecoverTableLine(line, pageHeight)).ToList())
            .Where(static band => band.Count > 0)
            .ToList();
        var tables = new List<StructuredTable>();
        // Leader-dominated bands should become leader tables, not generic band tables
        foreach (var band in bands) {
            if (band.Count == 0) continue;
            if (IsLeaderBand(band)) {
                var leader = BuildLeaderTableForBand(band);
                if (leader != null && leader.Rows.Count > 0) tables.Add(leader);
            }
        }
        // Then, attempt to form multi-band table groups with similar split positions (non-leader bands only)
        var nonLeaderBands = bands.Where(b => b.Count > 0 && !IsLeaderBand(b)).ToList();
        var grouped = DetectTablesAcrossBandGroups(nonLeaderBands);
        tables.AddRange(grouped);

        // Fallback per-band splits for remaining non-leader bands
        if (tables.Count == 0) {
            foreach (var band in nonLeaderBands) {
                var splits = InferSplits(band);
                if (splits.Count == 0) continue;
                var table = BuildTableFromLinesAndSplits(band, splits, "band-splits");
                if (table != null && table.Rows.Count >= 2 && HasValidatedRows(table, band)) tables.Add(table);
            }
        }
        List<TextLayoutEngine.TextLine> unmatchedLines = nonLeaderBands
            .SelectMany(static band => band)
            .Where(line => !IsCoveredByDetectedTable(line, tables))
            .Take(MaximumPositionedRecoveryLines)
            .ToList();
        List<StructuredTable> positionedTables = DetectPositionedCellTables(unmatchedLines, pageHeight);
        if (positionedTables.Count > 0) {
            tables.RemoveAll(table =>
                table.Rows.Count < 3 &&
                !string.Equals(table.Kind, "leaders", StringComparison.Ordinal) &&
                positionedTables.Any(positioned => IsSubsumedByPositionedTable(table, positioned)));
            tables.AddRange(positionedTables);
        }
        return tables;
    }

    private static bool IsSubsumedByPositionedTable(StructuredTable candidate, StructuredTable positioned) {
        if (candidate.Columns.Count == 0 || positioned.Columns.Count == 0 || positioned.Rows.Count <= candidate.Rows.Count) {
            return false;
        }

        double candidateLeft = candidate.Columns.Min(static column => Math.Min(column.From, column.To));
        double candidateRight = candidate.Columns.Max(static column => Math.Max(column.From, column.To));
        double positionedLeft = positioned.Columns.Min(static column => Math.Min(column.From, column.To));
        double positionedRight = positioned.Columns.Max(static column => Math.Max(column.From, column.To));
        double horizontalOverlap = Math.Max(0D, Math.Min(candidateRight, positionedRight) - Math.Max(candidateLeft, positionedLeft));
        double candidateWidth = candidateRight - candidateLeft;
        if (candidateWidth <= 0.001D || horizontalOverlap + 0.001D < candidateWidth * 0.5D) return false;

        double candidateTop = Math.Max(candidate.YTop, candidate.YBottom);
        double candidateBottom = Math.Min(candidate.YTop, candidate.YBottom);
        double positionedTop = Math.Max(positioned.YTop, positioned.YBottom);
        double positionedBottom = Math.Min(positioned.YTop, positioned.YBottom);
        double verticalOverlap = Math.Max(0D, Math.Min(candidateTop, positionedTop) - Math.Max(candidateBottom, positionedBottom));
        double candidateHeight = candidateTop - candidateBottom;
        if (candidateHeight > 0.001D && verticalOverlap + 0.001D < candidateHeight * 0.5D) return false;

        var positionedCells = new HashSet<string>(
            positioned.Rows.SelectMany(static row => row).Where(static cell => !string.IsNullOrWhiteSpace(cell)),
            StringComparer.Ordinal);
        return candidate.Rows
            .SelectMany(static row => row)
            .Where(static cell => !string.IsNullOrWhiteSpace(cell))
            .All(positionedCells.Contains);
    }

    private static bool IsCoveredByDetectedTable(
        TextLayoutEngine.TextLine line,
        List<StructuredTable> tables) {
        for (int index = 0; index < tables.Count; index++) {
            // Two-row band candidates are deliberately admitted only with strong
            // evidence, but they are still too weak to own the source geometry.
            // Let the independent positioned-cell pass inspect those lines so it
            // can recover a complete header/body region or a side-by-side table.
            if (tables[index].Rows.Count < 3 &&
                !string.Equals(tables[index].Kind, "leaders", StringComparison.Ordinal)) {
                continue;
            }
            double top = Math.Max(tables[index].YTop, tables[index].YBottom);
            double bottom = Math.Min(tables[index].YTop, tables[index].YBottom);
            if (line.Y > top + 0.001D || line.Y < bottom - 0.001D || tables[index].Columns.Count == 0) {
                continue;
            }

            double left = tables[index].Columns.Min(static column => Math.Min(column.From, column.To));
            double right = tables[index].Columns.Max(static column => Math.Max(column.From, column.To));
            double lineLeft = Math.Min(line.XStart, line.XEnd);
            double lineRight = Math.Max(line.XStart, line.XEnd);
            double overlap = Math.Max(0D, Math.Min(lineRight, right) - Math.Max(lineLeft, left));
            double narrowerWidth = Math.Min(lineRight - lineLeft, right - left);
            if (narrowerWidth > 0.001D && overlap + 0.001D >= narrowerWidth * 0.5D) return true;
        }
        return false;
    }

    internal static List<StructuredTable> DetectPositionedCellTables(
        IReadOnlyList<TextLayoutEngine.TextLine> lines,
        double? pageHeight = null) {
        var result = new List<StructuredTable>();
        var group = new List<PositionedRow>();
        int inspectedLines = 0;
        int inspectedCells = 0;
        foreach (TextLayoutEngine.TextLine line in lines.OrderByDescending(static line => line.Y)) {
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
        double from = 0D;
        double to = 0D;
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
                cells.Add(new PositionedCell(from, to, builder.ToString().Trim()));
                builder.Clear();
            } else if (gap > 1D && builder.Length > 0 && builder[builder.Length - 1] != ' ') {
                builder.Append(' ');
            }

            if (builder.Length == 0) from = span.X;
            builder.Append(span.Text);
            to = span.X + Math.Max(0D, span.Advance);
            if (cells.Count == MaximumPositionedRecoveryColumns) return null;
        }

        if (builder.Length > 0) cells.Add(new PositionedCell(from, to, builder.ToString().Trim()));
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
        result.Add(table);
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
                if (HasManyDigits(value) ||
                    string.Equals(value, "yes", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(value, "no", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(value, "false", StringComparison.OrdinalIgnoreCase)) return true;
            }
        }
        return false;
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
        internal PositionedCell(double from, double to, string text) {
            From = from;
            To = to;
            Text = text;
        }
        internal double From { get; }
        internal double To { get; }
        internal string Text { get; }
    }

    private static bool IsLeaderBand(List<TextLayoutEngine.TextLine> band) {
        if (band.Count == 0) return false;
        int leaderLines = 0; int nonEmpty = 0;
        foreach (var ln in band) {
            if (string.IsNullOrWhiteSpace(ln.Text)) continue; nonEmpty++;
            if (TryLeaderRowFromLine(ln, out _, out _, out _)) { leaderLines++; continue; }
            bool hasLeaderSpan = ln.Spans.Any(s => IsLeaderSpan(s.Text) && s.Text.Length >= 3);
            bool looksLeader = LooksLeaderText(ln.Text);
            if (hasLeaderSpan || looksLeader) leaderLines++;
        }
        if (nonEmpty == 0) return false;
        // Consider leader band if we have at least 3 leader-like rows, or >=30% of lines
        return leaderLines >= 3 || (leaderLines * 10 >= nonEmpty * 3);
    }

    private static StructuredTable? BuildLeaderTableForBand(List<TextLayoutEngine.TextLine> band) {
        var rows = new List<string[]>();
        double leftMin = double.MaxValue, leftMax = double.MinValue;
        double rightMin = double.MaxValue, rightMax = double.MinValue;
        foreach (var ln in band) {
            if (TryLeaderRowFromLine(ln, out var row, out var left, out var right)) {
                rows.Add(row);
                leftMin = Math.Min(leftMin, left.From);
                leftMax = Math.Max(leftMax, left.To);
                rightMin = Math.Min(rightMin, right.From);
                rightMax = Math.Max(rightMax, right.To);
            }
        }
        if (rows.Count == 0) return null;
        var t = new StructuredTable { YTop = band[0].Y, YBottom = band[band.Count - 1].Y, Kind = "leaders" };
        t.Columns.Add(new StructuredTableColumn { From = leftMin, To = leftMax });
        t.Columns.Add(new StructuredTableColumn { From = rightMin, To = rightMax });
        t.Rows.AddRange(rows);
        return t;
    }

    private static List<StructuredTable> DetectTablesAcrossBandGroups(List<List<TextLayoutEngine.TextLine>> bands) {
        var result = new List<StructuredTable>();
        // Pre-compute splits per band
        var bandSplits = new List<(int idx, List<TextLayoutEngine.TextLine> lines, List<double> splits)>();
        for (int i = 0; i < bands.Count; i++) {
            var b = bands[i]; if (b.Count == 0) continue;
            var sp = InferSplits(b);
            if (sp.Count == 0) continue;
            bandSplits.Add((i, b, sp));
        }
        int k = 0;
        while (k < bandSplits.Count) {
            int start = k;
            var baseSplits = bandSplits[k].splits;
            int end = k;
            var includedBridgeBandIndexes = new HashSet<int>();
            bool requiresAlignedCellSplits = false;
            // Extend while splits remain similar. An intervening band without
            // detectable splits is either retained as a strongly evidenced
            // spanning row or skipped when it belongs to an adjacent region.
            while (end + 1 < bandSplits.Count) {
                (int idx, List<TextLayoutEngine.TextLine> lines, List<double> splits) current = bandSplits[end];
                (int idx, List<TextLayoutEngine.TextLine> lines, List<double> splits) next = bandSplits[end + 1];
                bool hasNonLeftAlignedCells = BandsHaveNonLeftAlignedCells(current.lines, next.lines);
                if (next.idx > current.idx + 2 ||
                    (!AreSplitsSimilar(baseSplits, next.splits) && !hasNonLeftAlignedCells)) {
                    break;
                }
                requiresAlignedCellSplits |= hasNonLeftAlignedCells;

                InterveningBandDecision bridgeDecision = ClassifyInterveningBand(
                    bands,
                    current.idx,
                    next.idx,
                    baseSplits);
                if (bridgeDecision == InterveningBandDecision.Reject) {
                    break;
                }
                if (bridgeDecision == InterveningBandDecision.Include) {
                    includedBridgeBandIndexes.Add(current.idx + 1);
                }
                end++;
            }
            // Build table for [start..end], including a compatible header-only band immediately above it.
            var groupLines = new List<TextLayoutEngine.TextLine>();
            List<TextLayoutEngine.TextLine>? headerLines = TryGetPrecedingHeaderLines(
                bands,
                bandSplits[start].idx,
                baseSplits);
            if (headerLines is not null) {
                groupLines.AddRange(headerLines);
            }

            var splitBandIndexes = new HashSet<int>();
            for (int splitIndex = start; splitIndex <= end; splitIndex++) {
                splitBandIndexes.Add(bandSplits[splitIndex].idx);
            }
            for (int bandIndex = bandSplits[start].idx; bandIndex <= bandSplits[end].idx; bandIndex++) {
                if (splitBandIndexes.Contains(bandIndex) || includedBridgeBandIndexes.Contains(bandIndex)) {
                    groupLines.AddRange(bands[bandIndex]);
                }
            }
            List<double> effectiveSplits = requiresAlignedCellSplits
                ? InferAlignedCellSplits(groupLines, baseSplits.Count + 1) ?? baseSplits
                : baseSplits;
            var table = BuildTableFromLinesAndSplits(groupLines, effectiveSplits, "band-group");
            if (table != null &&
                (table.Rows.Count >= 3 || HasStrongTwoRowEvidence(table, groupLines)) &&
                HasValidatedRows(table, groupLines)) result.Add(table);
            k = end + 1;
        }
        return result;
    }

    private enum InterveningBandDecision {
        Reject,
        Include,
        Skip
    }

    private static InterveningBandDecision ClassifyInterveningBand(
        List<List<TextLayoutEngine.TextLine>> bands,
        int currentBandIndex,
        int nextBandIndex,
        List<double> splits) {
        if (nextBandIndex == currentBandIndex + 1) return InterveningBandDecision.Skip;
        if (nextBandIndex != currentBandIndex + 2 || splits.Count == 0) {
            return InterveningBandDecision.Reject;
        }

        List<TextLayoutEngine.TextLine> intervening = bands[currentBandIndex + 1];
        if (intervening.Count != 1) return InterveningBandDecision.Reject;
        TextLayoutEngine.TextLine line = intervening[0];
        if (!HasCompatibleRowRhythm(bands[currentBandIndex], line, bands[nextBandIndex])) {
            return InterveningBandDecision.Reject;
        }
        if (!HasMeaningfulHorizontalOverlap(line, bands[currentBandIndex], bands[nextBandIndex])) {
            return InterveningBandDecision.Reject;
        }

        string[] cells = SplitBySplits(line, splits);
        if (cells.Count(static cell => !string.IsNullOrWhiteSpace(cell)) != 1) {
            return InterveningBandDecision.Reject;
        }
        string text = line.Text.Trim();
        return HasEmphasizedText(line) ||
               IsUppercaseSectionLabel(text) ||
               (CrossesColumnBoundary(line, splits) && HasSpanningRowQualifier(text))
            ? InterveningBandDecision.Include
            : InterveningBandDecision.Reject;
    }

    private static bool HasMeaningfulHorizontalOverlap(
        TextLayoutEngine.TextLine line,
        List<TextLayoutEngine.TextLine> previousBand,
        List<TextLayoutEngine.TextLine> nextBand) {
        double tableLeft = previousBand.Concat(nextBand).Min(static candidate => Math.Min(candidate.XStart, candidate.XEnd));
        double tableRight = previousBand.Concat(nextBand).Max(static candidate => Math.Max(candidate.XStart, candidate.XEnd));
        double lineLeft = Math.Min(line.XStart, line.XEnd);
        double lineRight = Math.Max(line.XStart, line.XEnd);
        double overlap = Math.Max(0D, Math.Min(lineRight, tableRight) - Math.Max(lineLeft, tableLeft));
        double narrowerWidth = Math.Min(lineRight - lineLeft, tableRight - tableLeft);
        return narrowerWidth > 0.001D && overlap + 0.001D >= narrowerWidth * 0.5D;
    }

    private static bool CrossesColumnBoundary(TextLayoutEngine.TextLine line, IReadOnlyList<double> splits) {
        double left = Math.Min(line.XStart, line.XEnd);
        double right = Math.Max(line.XStart, line.XEnd);
        return splits.Any(split => left < split - 1D && right > split + 1D);
    }

    private static bool HasSpanningRowQualifier(string text) {
        if (text.Length is 0 or > 80) return false;
        string firstWord = text.Split((char[]?)null, 2, StringSplitOptions.RemoveEmptyEntries)[0].TrimEnd(':');
        return firstWord.Equals("amount", StringComparison.OrdinalIgnoreCase) ||
               firstWord.Equals("amounts", StringComparison.OrdinalIgnoreCase) ||
               firstWord.Equals("figure", StringComparison.OrdinalIgnoreCase) ||
               firstWord.Equals("figures", StringComparison.OrdinalIgnoreCase) ||
               firstWord.Equals("note", StringComparison.OrdinalIgnoreCase) ||
               firstWord.Equals("notes", StringComparison.OrdinalIgnoreCase) ||
               firstWord.Equals("subtotal", StringComparison.OrdinalIgnoreCase) ||
               firstWord.Equals("total", StringComparison.OrdinalIgnoreCase) ||
               firstWord.Equals("totals", StringComparison.OrdinalIgnoreCase) ||
               firstWord.Equals("value", StringComparison.OrdinalIgnoreCase) ||
               firstWord.Equals("values", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsUppercaseSectionLabel(string text) {
        bool hasLetter = false;
        for (int index = 0; index < text.Length; index++) {
            char value = text[index];
            if (!char.IsLetter(value)) continue;
            hasLetter = true;
            if (char.IsLower(value)) return false;
        }
        return hasLetter;
    }

    private static bool BandsHaveAlignedCells(
        List<TextLayoutEngine.TextLine> firstBand,
        List<TextLayoutEngine.TextLine> secondBand) {
        if (firstBand.Count != 1 || secondBand.Count != 1) return false;
        PositionedRow? first = TryCreatePositionedRow(firstBand[0]);
        PositionedRow? second = TryCreatePositionedRow(secondBand[0]);
        return first != null && second != null && PositionedRowsAlign(first, second);
    }

    private static bool BandsHaveNonLeftAlignedCells(
        List<TextLayoutEngine.TextLine> firstBand,
        List<TextLayoutEngine.TextLine> secondBand) {
        if (firstBand.Count != 1 || secondBand.Count != 1) return false;
        PositionedRow? first = TryCreatePositionedRow(firstBand[0]);
        PositionedRow? second = TryCreatePositionedRow(secondBand[0]);
        if (first == null || second == null || first.Cells.Count != second.Cells.Count) return false;

        bool hasNonLeftAlignment = false;
        for (int index = 0; index < first.Cells.Count; index++) {
            PositionedCell expected = first.Cells[index];
            PositionedCell current = second.Cells[index];
            bool leftAligned = Math.Abs(expected.From - current.From) <= 16D;
            bool centerAligned = Math.Abs(
                (expected.From + expected.To) / 2D -
                (current.From + current.To) / 2D) <= 16D;
            bool rightAligned = Math.Abs(expected.To - current.To) <= 16D;
            if (!leftAligned && !centerAligned && !rightAligned) return false;
            if (!leftAligned && (centerAligned || rightAligned)) hasNonLeftAlignment = true;
        }
        return hasNonLeftAlignment;
    }

    private static List<double>? InferAlignedCellSplits(
        List<TextLayoutEngine.TextLine> lines,
        int expectedColumnCount) {
        var rows = new List<PositionedRow>();
        for (int index = 0; index < lines.Count; index++) {
            PositionedRow? row = TryCreatePositionedRow(lines[index]);
            if (row != null && row.Cells.Count == expectedColumnCount) rows.Add(row);
        }
        if (rows.Count < 2) return null;

        var splits = new List<double>(expectedColumnCount - 1);
        for (int columnIndex = 0; columnIndex < expectedColumnCount - 1; columnIndex++) {
            double leftEdge = rows.Max(row => row.Cells[columnIndex].To);
            double rightEdge = rows.Min(row => row.Cells[columnIndex + 1].From);
            if (rightEdge <= leftEdge + 1D) return null;
            splits.Add(leftEdge + (rightEdge - leftEdge) / 2D);
        }
        return splits;
    }

    private static bool HasCompatibleRowRhythm(
        List<TextLayoutEngine.TextLine> previousBand,
        TextLayoutEngine.TextLine intervening,
        List<TextLayoutEngine.TextLine> nextBand) {
        double previousY = previousBand.Average(static line => line.Y);
        double nextY = nextBand.Average(static line => line.Y);
        double upperGap = previousY - intervening.Y;
        double lowerGap = intervening.Y - nextY;
        if (upperGap <= 0D || lowerGap <= 0D) return false;
        double smaller = Math.Min(upperGap, lowerGap);
        double larger = Math.Max(upperGap, lowerGap);
        return larger <= 36D && larger <= smaller * 1.75D;
    }

    private static bool IsCompactNonNarrativeRow(string text) {
        if (text.Length is 0 or > 80 ||
            text[text.Length - 1] is '.' or '!' or '?' ||
            text.StartsWith("Figure ", StringComparison.OrdinalIgnoreCase) ||
            text.StartsWith("Fig. ", StringComparison.OrdinalIgnoreCase) ||
            text.StartsWith("Table ", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }
        return text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length <= 8;
    }

    private static bool HasEmphasizedText(TextLayoutEngine.TextLine line) {
        PdfTextSpan[] spans = line.Spans
            .Where(static span => !string.IsNullOrWhiteSpace(span.Text))
            .ToArray();
        return spans.Length > 0 && spans.All(static span => IsEmphasizedFont(span.BaseFont));
    }

    private static List<TextLayoutEngine.TextLine>? TryGetPrecedingHeaderLines(
        List<List<TextLayoutEngine.TextLine>> bands,
        int bodyBandIndex,
        List<double> bodySplits) {
        int headerBandIndex = bodyBandIndex - 1;
        if (headerBandIndex < 0 || bodySplits.Count == 0) {
            return null;
        }

        List<TextLayoutEngine.TextLine> headerBand = bands[headerBandIndex];
        if (headerBand.Count != 1 || IsLeaderBand(headerBand)) {
            return null;
        }

        if (!IsCompactNonNarrativeRow(headerBand[0].Text.Trim()) ||
            (!BandsHaveAlignedCells(headerBand, bands[bodyBandIndex]) &&
             !HasEmphasizedText(headerBand[0]))) {
            return null;
        }

        string[] headerCells = SplitBySplits(headerBand[0], bodySplits);
        if (!LooksLikeHeaderRow(headerCells)) {
            return null;
        }

        return headerBand;
    }

    private static bool LooksLikeHeaderRow(string[] cells) {
        if (cells.Length < 2) {
            return false;
        }

        for (int i = 0; i < cells.Length; i++) {
            string cell = ContentStructureExtractor.NormalizeShattered(cells[i]).Trim();
            if (cell.Length == 0 ||
                (!cell.Any(char.IsLetter) && !cell.All(char.IsDigit))) {
                return false;
            }
        }

        return true;
    }

    private static bool AreSplitsSimilar(List<double> a, List<double> b) {
        if (a.Count != b.Count) return false;
        double tol = 16.0; // points
        for (int i = 0; i < a.Count; i++) if (Math.Abs(a[i] - b[i]) > tol) return false;
        return true;
    }

    private static StructuredTable? BuildTableFromLinesAndSplits(List<TextLayoutEngine.TextLine> lines, List<double> splits, string kind) {
        if (splits.Count == 0) return null;
        double minX = double.MaxValue, maxX = double.MinValue;
        foreach (var ln in lines) { minX = Math.Min(minX, ln.XStart); maxX = Math.Max(maxX, ln.XEnd); }
        var table = new StructuredTable {
            YTop = lines[0].Y,
            YBottom = lines[lines.Count - 1].Y,
            Kind = kind
        };
        double prev = minX;
        for (int i = 0; i <= splits.Count; i++) {
            double next = (i < splits.Count) ? splits[i] : maxX;
            table.Columns.Add(new StructuredTableColumn { From = prev, To = next });
            prev = next;
        }
        int cols = table.Columns.Count;
        foreach (var ln in lines) {
            var cells = SplitBySplits(ln, splits);
            if (cells.Length != cols) continue;
            bool anyContent = false; for (int i = 0; i < cells.Length; i++) if (!string.IsNullOrWhiteSpace(cells[i])) { anyContent = true; break; }
            if (!anyContent) continue;
            table.Rows.Add(cells);
        }
        return table.Rows.Count > 0 ? table : null;
    }

    private static bool LooksLikeCompactHeaderRow(string[] cells) {
        if (!LooksLikeHeaderRow(cells)) return false;
        int words = 0;
        for (int index = 0; index < cells.Length; index++) {
            string value = ContentStructureExtractor.NormalizeShattered(cells[index]).Trim();
            int cellWords = value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
            if (cellWords > 4 || value[value.Length - 1] is '.' or '!' or '?') return false;
            words += cellWords;
        }
        return words <= cells.Length * 3;
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
        bool compactGrid = HasCompactCellGrid(table) && !HasPageColumnLikeGutters(sourceLines);
        return LooksLikeSparseFormGrid(table) ||
               (dense && (
                   hasTabularValueEvidence ||
                   HasEmphasizedHeader(sourceLines) ||
                   compactGrid ||
                   HasStableColumnAnchors(table, sourceLines)));
    }

    private static bool LooksLikeSparseFormGrid(StructuredTable table) {
        if (table.Rows.Count < 3 || !LooksLikeCompactHeaderRow(table.Rows[0])) return false;

        int sparseLabelRows = 0;
        for (int rowIndex = 1; rowIndex < table.Rows.Count; rowIndex++) {
            string[] row = table.Rows[rowIndex];
            int populatedCells = 0;
            int populatedColumn = -1;
            for (int columnIndex = 0; columnIndex < Math.Min(row.Length, table.Columns.Count); columnIndex++) {
                if (string.IsNullOrWhiteSpace(row[columnIndex])) continue;
                populatedCells++;
                populatedColumn = columnIndex;
            }
            if (populatedCells != 1 || populatedColumn != 0) continue;

            string label = ContentStructureExtractor.NormalizeShattered(row[0]).Trim();
            int words = label.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
            if (words is >= 1 and <= 5 && label[label.Length - 1] is not ('.' or '!' or '?')) {
                sparseLabelRows++;
            }
        }
        return sparseLabelRows >= 2;
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

    private static bool HasStrongTwoRowEvidence(
        StructuredTable table,
        List<TextLayoutEngine.TextLine> sourceLines) {
        if (table.Rows.Count != 2 || sourceLines.Count != 2) return false;
        if (!LooksLikeHeaderRow(table.Rows[0])) return false;
        return table.Rows[1].Any(IsTabularValue) ||
               (HasStableColumnAnchors(table, sourceLines) && HasEmphasizedHeader(sourceLines));
    }

    private static bool HasCompactCellGrid(StructuredTable table) {
        if (table.Rows.Count < 2 || !LooksLikeHeaderRow(table.Rows[0])) return false;
        int populatedCells = 0;
        int wordCount = 0;
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            string[] row = table.Rows[rowIndex];
            if (row.Length < table.Columns.Count) return false;
            for (int columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++) {
                string value = row[columnIndex].Trim();
                if (value.Length == 0) return false;
                populatedCells++;
                wordCount += value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
            }
        }
        return populatedCells > 0 && wordCount <= populatedCells * 2;
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
        return headerSpans.Length >= 2 && headerSpans.All(static span => IsEmphasizedFont(span.BaseFont));
    }

    private static bool IsEmphasizedFont(string? baseFont) =>
        baseFont?.IndexOf("Bold", StringComparison.OrdinalIgnoreCase) >= 0 ||
        baseFont?.IndexOf("Black", StringComparison.OrdinalIgnoreCase) >= 0 ||
        baseFont?.IndexOf("Demi", StringComparison.OrdinalIgnoreCase) >= 0 ||
        baseFont?.IndexOf("SemiBold", StringComparison.OrdinalIgnoreCase) >= 0;

    private static bool IsTabularValue(string value) =>
        HasManyDigits(value) ||
        string.Equals(value, "yes", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "no", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "false", StringComparison.OrdinalIgnoreCase);

    public static StructuredTable? DetectLeaderTable(
        List<TextLayoutEngine.TextLine> lines,
        double? pageHeight = null) {
        var candidates = lines
            .Where(line => CanRecoverTableLine(line, pageHeight))
            .Where(static line => !string.IsNullOrWhiteSpace(line.Text))
            .ToList();
        if (candidates.Count == 0) return null;
        var rows = new List<string[]>();
        double leftMin = double.MaxValue, leftMax = double.MinValue;
        double rightMin = double.MaxValue, rightMax = double.MinValue;
        foreach (var ln in candidates) {
            if (TryLeaderRowFromLine(ln, out var row, out var leftBounds, out var rightBounds)) {
                rows.Add(row);
                leftMin = Math.Min(leftMin, leftBounds.From);
                leftMax = Math.Max(leftMax, leftBounds.To);
                rightMin = Math.Min(rightMin, rightBounds.From);
                rightMax = Math.Max(rightMax, rightBounds.To);
            }
        }
        if (rows.Count < 2) return null;
        var table = new StructuredTable {
            YTop = candidates[0].Y,
            YBottom = candidates[candidates.Count - 1].Y,
            Kind = "leaders"
        };
        table.Columns.Add(new StructuredTableColumn { From = leftMin, To = leftMax });
        table.Columns.Add(new StructuredTableColumn { From = rightMin, To = rightMax });
        table.Rows.AddRange(rows);
        return table;
    }

    /// <summary>
    /// Band-aware detection that first infers stable column split positions within each band,
    /// then splits lines consistently into those columns.
    /// </summary>
    public static List<string[]> DetectFromBands(
        List<List<TextLayoutEngine.TextLine>> bands,
        double? pageHeight = null) {
        var all = new List<string[]>();
        foreach (var band in bands) {
            List<TextLayoutEngine.TextLine> recoverable = band
                .Where(line => CanRecoverTableLine(line, pageHeight))
                .ToList();
            if (recoverable.Count == 0) continue;
            var splits = InferSplits(recoverable);
            if (splits.Count == 0) {
                // fallback to per-line gap splitting
                foreach (var ln in recoverable) {
                    if (ln.Spans.Count < 2) continue;
                    var cells = SplitByGaps(ln);
                    if (cells.Length >= 2 && LooksTabular(cells)) all.Add(cells);
                }
                continue;
            }
            // Consistent splitting using inferred splits
            int cols = splits.Count + 1;
            foreach (var ln in recoverable) {
                var cells = SplitBySplits(ln, splits);
                if (cells.Length == cols) {
                    bool any = false; for (int i = 0; i < cells.Length; i++) if (!string.IsNullOrWhiteSpace(cells[i])) { any = true; break; }
                    if (any) all.Add(cells);
                }
            }
        }
        return all;
    }

    private static bool CanRecoverTableLine(TextLayoutEngine.TextLine line, double? pageHeight) =>
        line.Spans.Count > 0 &&
        line.Spans.All(span => span.CanProjectCompleteText(pageHeight));

    private static List<double> InferSplits(List<TextLayoutEngine.TextLine> lines) {
        // Collect candidate split X positions as midpoints of large gaps between adjacent spans
        var cands = new List<double>();
        int eligibleLines = 0;
        foreach (var ln in lines) {
            if (ln.Spans.Count < 2) continue;
            eligibleLines++;
            // Dot-leader spans are strong split hints
            for (int k = 0; k < ln.Spans.Count; k++) {
                var s = ln.Spans[k];
                if (IsLeaderSpan(s.Text)) {
                    double mid = s.X + Math.Max(0, s.Advance) / 2.0;
                    cands.Add(mid);
                }
            }
            for (int i = 1; i < ln.Spans.Count; i++) {
                var prev = ln.Spans[i - 1]; var curSpan = ln.Spans[i];
                double prevEnd = prev.X + Math.Max(0, prev.Advance);
                double gap = curSpan.X - prevEnd;
                double em = Math.Max(prev.FontSize, curSpan.FontSize);
                double threshold = Math.Max(18.0, em * 2.0);
                if (gap >= threshold) {
                    double mid = prevEnd + (gap / 2.0);
                    cands.Add(mid);
                }
            }
        }
        if (eligibleLines == 0 || cands.Count == 0) return new List<double>();
        // Histogram candidates into 4pt bins and select peaks with sufficient votes
        double binW = 4.0;
        double minX = cands.Min(); double maxX = cands.Max();
        int bins = Math.Max(1, (int)Math.Ceiling((maxX - minX) / binW));
        var hist = new int[bins];
        foreach (var x in cands) {
            int b = (int)Math.Floor((x - minX) / binW);
            if (b < 0) b = 0; if (b >= bins) b = bins - 1; hist[b]++;
        }
        int voteCut = eligibleLines == 1 ? 1 : Math.Max(2, (int)Math.Ceiling(eligibleLines * 0.35));
        var peaks = new List<double>();
        for (int b = 0; b < bins; b++) if (hist[b] >= voteCut) peaks.Add(minX + b * binW + binW / 2.0);
        if (peaks.Count == 0) {
            // Fallback for narrow bands: pick the strongest bin if any votes exist
            int maxVotes = 0; int maxBin = -1;
            for (int b = 0; b < bins; b++) if (hist[b] > maxVotes) { maxVotes = hist[b]; maxBin = b; }
            if (maxVotes > 0 && maxBin >= 0) peaks.Add(minX + maxBin * binW + binW / 2.0);
            else return new List<double>();
        }
        // Merge nearby peaks (< 16pt apart)
        peaks.Sort();
        var merged = new List<double>();
        double acc = peaks[0]; int count = 1;
        for (int i = 1; i < peaks.Count; i++) {
            if (Math.Abs(peaks[i] - acc) < 16.0) { acc = (acc * count + peaks[i]) / (count + 1); count++; }
            else { merged.Add(acc); acc = peaks[i]; count = 1; }
        }
        merged.Add(acc);
        // Limit to a reasonable number of splits to avoid over-fragmentation
        if (merged.Count > 6) merged = merged.Take(6).ToList();
        return merged;
    }

    private static string[] SplitBySplits(TextLayoutEngine.TextLine ln, List<double> splits) {
        int cols = splits.Count + 1;
        var cellBuilders = new System.Text.StringBuilder[cols];
        for (int i = 0; i < cols; i++) cellBuilders[i] = new System.Text.StringBuilder();
        int ColIndex(double x) { int idx = 0; while (idx < splits.Count && x >= splits[idx]) idx++; return idx; }
        for (int i = 0; i < ln.Spans.Count; i++) {
            var s = ln.Spans[i];
            int cidx = ColIndex(s.X);
            var sb = cellBuilders[cidx];
            if (sb.Length > 0 && !char.IsWhiteSpace(sb[sb.Length - 1])) sb.Append(' ');
            sb.Append(s.Text);
        }
        var cells = new string[cols];
        for (int i = 0; i < cols; i++) cells[i] = cellBuilders[i].ToString().Trim();
        return cells;
    }

    private static string[] SplitByGaps(TextLayoutEngine.TextLine ln) {
        // Determine gaps between spans using XEnd(prev) -> XStart(next)
        double ThresholdFor(PdfTextSpan prev, PdfTextSpan next) {
            double em = Math.Max(prev.FontSize, next.FontSize);
            return Math.Max(18.0, em * 2.0); // 18pt or 2em
        }
        var cells = new List<string>();
        var current = new System.Text.StringBuilder();
        for (int i = 0; i < ln.Spans.Count; i++) {
            var s = ln.Spans[i];
            if (i > 0) {
                var p = ln.Spans[i - 1];
                double prevEnd = p.X + Math.Max(0, p.Advance);
                double gap = s.X - prevEnd;
                if (gap > ThresholdFor(p, s)) {
                    // split to a new cell
                    cells.Add(current.ToString().Trim());
                    current.Clear();
                } else if (gap > 1.0 && (current.Length > 0 && current[current.Length - 1] != ' ')) {
                    // small gap -> ensure single space
                    current.Append(' ');
                }
            }
            current.Append(s.Text);
        }
        if (current.Length > 0) cells.Add(current.ToString().Trim());
        return cells.ToArray();
    }

    private static bool LooksTabular(string[] cells) {
        // Require at least one numeric-ish cell and avoid one-word rows
        bool anyNumeric = cells.Any(c => HasManyDigits(c));
        bool hasContent = cells.Any(c => c.Length >= 2);
        return anyNumeric && hasContent;
    }

    private static bool HasManyDigits(string s) {
        int digits = 0; for (int i = 0; i < s.Length; i++) if (char.IsDigit(s[i])) digits++;
        return digits >= Math.Max(2, s.Length / 4);
    }

    private static bool IsLeaderSpan(string s) {
        if (string.IsNullOrEmpty(s)) return false;
        char c = s[0];
        if (c != '.' && c != '-' && c != '_') return false;
        for (int i = 1; i < s.Length; i++) if (s[i] != c) return false; return true;
    }

    private static bool LooksLeaderText(string s) {
        if (string.IsNullOrWhiteSpace(s)) return false;
        char previous = '\0';
        int runLength = 0;
        for (int i = 0; i < s.Length; i++) {
            char current = s[i];
            if (current != '.' && current != '-' && current != '_') {
                previous = '\0';
                runLength = 0;
                continue;
            }

            runLength = current == previous ? runLength + 1 : 1;
            if (runLength >= 3) return true;
            previous = current;
        }
        return false;
    }

    private static bool TryLeaderRowFromLine(TextLayoutEngine.TextLine ln, out string[] row, out (double From,double To) left, out (double From,double To) right) {
        row = Array.Empty<string>(); left = (0,0); right=(0,0);
        // Find a leader span in this line
        int leaderIdx = -1;
        for (int i = 0; i < ln.Spans.Count; i++) if (IsLeaderSpan(ln.Spans[i].Text) && ln.Spans[i].Text.Length >= 3) { leaderIdx = i; break; }
        if (leaderIdx < 0) return false;
        // Left label: join spans before leader (preserve minimal spaces)
        var sbLeft = new System.Text.StringBuilder();
        double leftFrom = double.MaxValue, leftTo = double.MinValue;
        for (int i = 0; i < leaderIdx; i++) {
            var s = ln.Spans[i];
            if (sbLeft.Length > 0) sbLeft.Append(' ');
            sbLeft.Append(s.Text);
            leftFrom = Math.Min(leftFrom, s.X);
            leftTo = Math.Max(leftTo, s.X + Math.Max(0, s.Advance));
        }
        string leftText = CleanLeftLabel(sbLeft.ToString());
        // Right value: consume the value spans after leader, preserving numeric punctuation.
        var sbRight = new System.Text.StringBuilder();
        double rightFrom = double.MaxValue, rightTo = double.MinValue;
        for (int i = leaderIdx + 1; i < ln.Spans.Count; i++) {
            var s = ln.Spans[i];
            if (IsLeaderSpan(s.Text)) {
                continue;
            }

            if (sbRight.Length > 0 && sbRight[sbRight.Length - 1] != ' ') sbRight.Append(' ');
            sbRight.Append(s.Text);
            rightFrom = Math.Min(rightFrom, s.X);
            rightTo = Math.Max(rightTo, s.X + Math.Max(0, s.Advance));
        }
        string rightText = NormalizeLeaderValue(sbRight.ToString());
        // Sanity checks
        if (leftText.Length == 0 || rightText.Length == 0) return false;
        row = new [] { leftText, rightText };
        left = (leftFrom, leftTo);
        right = (rightFrom, rightTo);
        return true;
    }

    private static string CleanLeftLabel(string s) {
        if (string.IsNullOrEmpty(s)) return s;
        // Normalize spaces
        s = System.Text.RegularExpressions.Regex.Replace(s, "\\s+", " ").Trim();
        // Remove trailing leader characters on label
        s = s.Trim('.', '-', '_');
        // Remove repeated dot groups inside label
        s = System.Text.RegularExpressions.Regex.Replace(s, "[.]{2,}", ".");
        // Tidy quotes and parentheses spacing
        s = s.Replace(" ' ", " '").Replace("( ", "(").Replace(" )", ")");
        // Re-insert spaces around common glued prepositions if camel-cased inside
        s = System.Text.RegularExpressions.Regex.Replace(s, "([A-Za-z])of([A-Z])", "$1 of $2");
        s = System.Text.RegularExpressions.Regex.Replace(s, "([a-z]{2,})of([A-Z])", "$1 of $2");
        s = System.Text.RegularExpressions.Regex.Replace(s, "([A-Za-z])in([A-Z])", "$1 in $2");
        s = System.Text.RegularExpressions.Regex.Replace(s, "([a-z]{2,})in([A-Z])", "$1 in $2");
        s = System.Text.RegularExpressions.Regex.Replace(s, "([A-Za-z])and([A-Z])", "$1 and $2");
        s = System.Text.RegularExpressions.Regex.Replace(s, "([a-z]{2,})and([A-Z])", "$1 and $2");
        // generic lower->Upper split (camel-case -> spaced)
        s = System.Text.RegularExpressions.Regex.Replace(s, "([a-z])([A-Z])", "$1 $2");
        // Collapse micro-token shattering (aggressive but safe-ish for leaders)
        var parts = s.Split(' ');
        if (parts.Length <= 2) return s;
        bool Wordish(string t) { for (int i = 0; i < t.Length; i++) { char c = t[i]; if (!(char.IsLetterOrDigit(c) || c=='\''||c=='-'||c=='/')) return false; } return t.Length>0; }
        bool ShortAbbrev(string t) { if (t.Length==0 || t.Length>3) return false; for (int i=0;i<t.Length;i++) if(!char.IsUpper(t[i])) return false; return true; }
        int shortCount = parts.Count(p => p.Length <= 2 && Wordish(p));
        if (!(shortCount >= 2 || shortCount * 4 >= parts.Length)) return s;
        var sb = new System.Text.StringBuilder(s.Length);
        sb.Append(parts[0]);
        for (int i = 1; i < parts.Length; i++) {
            string prev = parts[i-1]; string cur = parts[i];
            bool joinSmall = Wordish(prev) && Wordish(cur) && !ShortAbbrev(prev) && !ShortAbbrev(cur) && (prev.Length<=2 || cur.Length<=2);
            bool nextShort = (i+1<parts.Length) && parts[i+1].Length<=2 && Wordish(parts[i+1]) && !ShortAbbrev(parts[i+1]);
            if (joinSmall || (Wordish(cur)&&cur.Length<=2 && nextShort)) sb.Append(cur);
            else sb.Append(' ').Append(cur);
        }
        return sb.ToString().Replace("  ", " ");
    }

    private static string NormalizeLeaderValue(string value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return string.Empty;
        }

        string normalized = System.Text.RegularExpressions.Regex.Replace(value.Trim(), "\\s+", " ");
        normalized = System.Text.RegularExpressions.Regex.Replace(normalized, "\\s*([.,])\\s*", "$1");
        normalized = System.Text.RegularExpressions.Regex.Replace(normalized, "([$€£])\\s+", "$1");
        normalized = normalized.Trim('.');

        bool hasDigit = false;
        for (int i = 0; i < normalized.Length; i++) {
            if (char.IsDigit(normalized[i])) {
                hasDigit = true;
                break;
            }
        }

        return hasDigit ? normalized : string.Empty;
    }
}

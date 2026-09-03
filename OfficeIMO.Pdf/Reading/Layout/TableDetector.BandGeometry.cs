using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Pdf;

internal static partial class TableDetector {
    private static bool BandsHaveCompatibleVerticalRhythm(
        List<TextLayoutEngine.TextLine> previousBand,
        List<TextLayoutEngine.TextLine> currentBand,
        List<TextLayoutEngine.TextLine> nextBand) {
        if (previousBand.Count != 1 || currentBand.Count != 1 || nextBand.Count != 1) return false;
        double previousGap = previousBand[0].Y - currentBand[0].Y;
        double nextGap = currentBand[0].Y - nextBand[0].Y;
        if (previousGap <= 0D || nextGap <= 0D) return false;
        double smallerGap = Math.Min(previousGap, nextGap);
        double largerGap = Math.Max(previousGap, nextGap);
        return largerGap <= smallerGap * 1.75D;
    }

    private static bool BandsAlignUsingSplits(
        List<TextLayoutEngine.TextLine> firstBand,
        List<TextLayoutEngine.TextLine> secondBand,
        List<double> splits) {
        if (firstBand.Count != 1 || secondBand.Count != 1 || splits.Count == 0) return false;
        List<(double From, double To)>? firstCells = GetSplitCellBounds(firstBand[0], splits);
        List<(double From, double To)>? secondCells = GetSplitCellBounds(secondBand[0], splits);
        if (firstCells is null || secondCells is null || firstCells.Count != secondCells.Count) return false;

        for (int index = 0; index < firstCells.Count; index++) {
            (double firstFrom, double firstTo) = firstCells[index];
            (double secondFrom, double secondTo) = secondCells[index];
            bool leftAligned = Math.Abs(firstFrom - secondFrom) <= 16D;
            bool centerAligned = Math.Abs(
                (firstFrom + firstTo) / 2D -
                (secondFrom + secondTo) / 2D) <= 16D;
            bool rightAligned = Math.Abs(firstTo - secondTo) <= 16D;
            if (!leftAligned && !centerAligned && !rightAligned) return false;
        }
        return true;
    }

    private static List<(double From, double To)>? GetSplitCellBounds(
        TextLayoutEngine.TextLine line,
        List<double> splits) {
        int columnCount = splits.Count + 1;
        var from = Enumerable.Repeat(double.PositiveInfinity, columnCount).ToArray();
        var to = Enumerable.Repeat(double.NegativeInfinity, columnCount).ToArray();
        for (int spanIndex = 0; spanIndex < line.Spans.Count; spanIndex++) {
            PdfTextSpan span = line.Spans[spanIndex];
            if (string.IsNullOrWhiteSpace(span.Text)) continue;
            int columnIndex = 0;
            while (columnIndex < splits.Count && span.X >= splits[columnIndex]) columnIndex++;
            from[columnIndex] = Math.Min(from[columnIndex], span.X);
            to[columnIndex] = Math.Max(to[columnIndex], span.X + Math.Max(0D, span.Advance));
        }

        var cells = new List<(double From, double To)>(columnCount);
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            if (double.IsPositiveInfinity(from[columnIndex])) return null;
            cells.Add((from[columnIndex], to[columnIndex]));
        }
        return cells;
    }

    private static bool BandsContainAlignedCells(
        List<TextLayoutEngine.TextLine> firstBand,
        List<TextLayoutEngine.TextLine> secondBand) {
        var firstRowsByColumnCount = new Dictionary<int, PositionedRow>();
        for (int firstIndex = 0; firstIndex < firstBand.Count; firstIndex++) {
            PositionedRow? first = TryCreatePositionedRow(firstBand[firstIndex]);
            if (first is not null && !firstRowsByColumnCount.ContainsKey(first.Cells.Count)) {
                firstRowsByColumnCount.Add(first.Cells.Count, first);
            }
        }
        for (int secondIndex = 0; secondIndex < secondBand.Count; secondIndex++) {
            PositionedRow? second = TryCreatePositionedRow(secondBand[secondIndex]);
            if (second is not null &&
                firstRowsByColumnCount.TryGetValue(second.Cells.Count, out PositionedRow? first) &&
                PositionedRowsAlign(first, second)) {
                return true;
            }
        }
        return false;
    }

    private static bool SplitsSeparatePositionedCells(
        List<TextLayoutEngine.TextLine> lines,
        List<double> splits,
        int expectedColumnCount) {
        for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
            PositionedRow? row = TryCreatePositionedRow(lines[lineIndex]);
            if (row is null || row.Cells.Count != expectedColumnCount) continue;
            for (int boundaryIndex = 0; boundaryIndex < splits.Count; boundaryIndex++) {
                double split = splits[boundaryIndex];
                if (split <= row.Cells[boundaryIndex].LastSpanStart ||
                    split > row.Cells[boundaryIndex + 1].From) {
                    return false;
                }
            }
        }
        return true;
    }

    private static string[] MergeBandCellsBySplits(
        List<TextLayoutEngine.TextLine> band,
        List<double> splits) {
        var cells = new string[splits.Count + 1];
        for (int columnIndex = 0; columnIndex < cells.Length; columnIndex++) cells[columnIndex] = string.Empty;
        for (int lineIndex = 0; lineIndex < band.Count; lineIndex++) {
            string[] lineCells = SplitBySplits(band[lineIndex], splits);
            for (int columnIndex = 0; columnIndex < cells.Length; columnIndex++) {
                string value = lineCells[columnIndex].Trim();
                if (value.Length == 0) continue;
                cells[columnIndex] = string.IsNullOrEmpty(cells[columnIndex])
                    ? value
                    : cells[columnIndex] + " " + value;
            }
        }
        return cells;
    }
}

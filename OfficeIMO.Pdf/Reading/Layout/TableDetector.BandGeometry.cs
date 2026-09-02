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

    private static bool LooksLikeNaturalSpanningPhrase(TextLayoutEngine.TextLine line) {
        PdfTextSpan[] spans = line.Spans
            .Where(static span => !string.IsNullOrWhiteSpace(span.Text))
            .OrderBy(static span => span.X)
            .ToArray();
        if (spans.Length < 2 || spans.Any(static span => span.Text.Any(char.IsDigit))) return false;

        for (int index = 1; index < spans.Length; index++) {
            PdfTextSpan previous = spans[index - 1];
            PdfTextSpan current = spans[index];
            double gap = current.X - (previous.X + Math.Max(0D, previous.Advance));
            double wordSpacing = Math.Max(4D, Math.Max(previous.FontSize, current.FontSize) * 1.25D);
            if (gap < -1D || gap > wordSpacing) return false;
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
}

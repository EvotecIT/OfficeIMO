using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Pdf;

internal static partial class TableDetector {
    private sealed class AlignedSplitAccumulator {
        private readonly int _expectedColumnCount;
        private double[] _maximumInkEnds;
        private double[] _maximumLastSpanStarts;
        private double[] _minimumNextStarts;
        private int _rowCount;

        internal AlignedSplitAccumulator(
            int expectedColumnCount,
            IReadOnlyList<TextLayoutEngine.TextLine> initialLines) {
            _expectedColumnCount = expectedColumnCount;
            int boundaryCount = Math.Max(0, expectedColumnCount - 1);
            _maximumInkEnds = Enumerable.Repeat(double.NegativeInfinity, boundaryCount).ToArray();
            _maximumLastSpanStarts = Enumerable.Repeat(double.NegativeInfinity, boundaryCount).ToArray();
            _minimumNextStarts = Enumerable.Repeat(double.PositiveInfinity, boundaryCount).ToArray();
            TryAppend(initialLines, requireValidSplits: false);
        }

        internal bool TryAppend(
            IReadOnlyList<TextLayoutEngine.TextLine> lines,
            bool requireValidSplits) {
            double[] maximumInkEnds = (double[])_maximumInkEnds.Clone();
            double[] maximumLastSpanStarts = (double[])_maximumLastSpanStarts.Clone();
            double[] minimumNextStarts = (double[])_minimumNextStarts.Clone();
            int rowCount = _rowCount;

            for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
                PositionedRow? row = TryCreatePositionedRow(lines[lineIndex]);
                if (row == null || row.Cells.Count != _expectedColumnCount) continue;
                rowCount++;
                for (int columnIndex = 0; columnIndex < _expectedColumnCount - 1; columnIndex++) {
                    maximumInkEnds[columnIndex] = Math.Max(
                        maximumInkEnds[columnIndex],
                        row.Cells[columnIndex].To);
                    maximumLastSpanStarts[columnIndex] = Math.Max(
                        maximumLastSpanStarts[columnIndex],
                        row.Cells[columnIndex].LastSpanStart);
                    minimumNextStarts[columnIndex] = Math.Min(
                        minimumNextStarts[columnIndex],
                        row.Cells[columnIndex + 1].From);
                }
            }

            if (requireValidSplits &&
                !TryBuildSplits(
                    rowCount,
                    maximumInkEnds,
                    maximumLastSpanStarts,
                    minimumNextStarts,
                    out _)) {
                return false;
            }

            _rowCount = rowCount;
            _maximumInkEnds = maximumInkEnds;
            _maximumLastSpanStarts = maximumLastSpanStarts;
            _minimumNextStarts = minimumNextStarts;
            return true;
        }

        internal List<double>? GetSplits() => TryBuildSplits(
            _rowCount,
            _maximumInkEnds,
            _maximumLastSpanStarts,
            _minimumNextStarts,
            out List<double>? splits)
            ? splits
            : null;

        private static bool TryBuildSplits(
            int rowCount,
            double[] maximumInkEnds,
            double[] maximumLastSpanStarts,
            double[] minimumNextStarts,
            out List<double>? splits) {
            splits = null;
            if (rowCount < 2) return false;

            var candidate = new List<double>(maximumInkEnds.Length);
            for (int boundaryIndex = 0; boundaryIndex < maximumInkEnds.Length; boundaryIndex++) {
                double leftEdge = maximumInkEnds[boundaryIndex];
                double rightEdge = minimumNextStarts[boundaryIndex];
                if (rightEdge <= leftEdge + 1D) {
                    // Text ink may overlap the next column even though every span start remains
                    // separable. SplitBySplits assigns spans by their start coordinate, so this
                    // narrower boundary is still safe and retains prose-heavy table cells.
                    leftEdge = maximumLastSpanStarts[boundaryIndex];
                    if (rightEdge <= leftEdge + 1D) return false;
                }
                candidate.Add(leftEdge + (rightEdge - leftEdge) / 2D);
            }

            splits = candidate;
            return true;
        }
    }
}

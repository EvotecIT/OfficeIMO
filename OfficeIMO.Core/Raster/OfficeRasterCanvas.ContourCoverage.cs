using System;
using System.Collections.Generic;
#if NET8_0_OR_GREATER
using System.Buffers;
#endif

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private const int ContourSubScanlines = 8;
    private const long MaximumContourCrossingWork = 512_000_000L;
    private const long MaximumContourRowCrossingWork = 16_000_000L;
    private const int MaximumRetainedContourCrossings = 1_000_000;

    private void FillContours(IReadOnlyList<IReadOnlyList<OfficePoint>> contours, OfficeColor color, OfficeFillRule fillRule) {
        if (color.A == 0) return;
        FillContourPaint(contours, fillRule, (_, _) => color);
    }

    /// <summary>
    /// Accumulates horizontal area coverage before painting a pixel once. Vertex heights split
    /// the sub-scanlines, retaining thin horizontal details even between the regular sample rows.
    /// </summary>
    internal void FillContourPaint(IReadOnlyList<IReadOnlyList<OfficePoint>> contours, OfficeFillRule fillRule, Func<double, double, OfficeColor> paint,
        IReadOnlyList<IReadOnlyList<OfficePoint>>? unionContours = null) {
        if (contours == null || contours.Count == 0) return;
        var boundaries = new List<double>();
        double minX = double.PositiveInfinity, maxX = double.NegativeInfinity;
        var boundsContours = new List<IReadOnlyList<OfficePoint>>(contours);
        if (unionContours != null) boundsContours.AddRange(unionContours);
        long contourEdges = 0L;
        foreach (IReadOnlyList<OfficePoint> contour in boundsContours) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (contour.Count < 3) continue;
            contourEdges += contour.Count;
            foreach (OfficePoint point in contour) {
                if (!IsFinite(point.X) || !IsFinite(point.Y)) return;
                minX = Math.Min(minX, point.X);
                maxX = Math.Max(maxX, point.X);
                boundaries.Add(point.Y);
            }
        }
        if (boundaries.Count == 0 || maxX <= 0D || minX >= Width) return;
        boundaries.Sort();
        double minY = boundaries[0], maxY = boundaries[boundaries.Count - 1];
        if (maxY <= 0D || minY >= Height) return;
        int left = (int)Math.Max(0D, Math.Floor(minX));
        int right = (int)Math.Min(Width - 1D, Math.Ceiling(maxX) - 1D);
        int top = (int)Math.Max(0D, Math.Floor(minY));
        int bottom = (int)Math.Min(Height - 1D, Math.Ceiling(maxY) - 1D);
        if (right < left || bottom < top) return;
        int tileLength = Math.Min(right - left + 1, MaximumContourCoverageTileWidth);
#if NET8_0_OR_GREATER
        double[] coverage = ArrayPool<double>.Shared.Rent(tileLength);
#else
        double[] coverage = new double[tileLength];
#endif
        try {
            var rowBoundaries = new List<double>();
            var crossings = new List<ContourCrossing>();
            // Retain one bounded row buffer rather than copying every sub-scanline into
            // a new list. The ranges stay in scanline order, preserving coverage sums.
            var rowCrossings = new List<ContourCrossing>();
            var scanlines = new List<(double Weight, int Start, int Count)>();
            int boundaryIndex = 0;
            long crossingWork = 0L;
            for (int y = top; y <= bottom; y++) {
                _cancellationToken.ThrowIfCancellationRequested();
                rowBoundaries.Clear();
                for (int sample = 0; sample <= ContourSubScanlines; sample++) rowBoundaries.Add(y + sample / (double)ContourSubScanlines);
                while (boundaryIndex < boundaries.Count && boundaries[boundaryIndex] <= y) boundaryIndex++;
                while (boundaryIndex < boundaries.Count && boundaries[boundaryIndex] < y + 1D) rowBoundaries.Add(boundaries[boundaryIndex++]);
                rowBoundaries.Sort();
                long rowWork = contourEdges * (rowBoundaries.Count - 1L);
                if (rowWork > MaximumContourRowCrossingWork ||
                    rowWork > MaximumContourCrossingWork - crossingWork) {
                    throw new InvalidOperationException("Contour coverage work exceeds the rasterization limit.");
                }
                crossingWork += rowWork;
                scanlines.Clear();
                rowCrossings.Clear();
                int retainedCrossings = 0;
                for (int index = 1; index < rowBoundaries.Count; index++) {
                    double low = rowBoundaries[index - 1], high = rowBoundaries[index];
                    if (high <= low) continue;
                    crossings.Clear();
                    AddContourCrossings(contours, (low + high) / 2D, crossings);
                    if (unionContours != null) AddContourCrossings(unionContours, (low + high) / 2D, crossings, true);
                    crossings.Sort(ContourCrossingComparer.Instance);
                    if (crossings.Count >= 2) {
                        if (crossings.Count > MaximumRetainedContourCrossings - retainedCrossings) {
                            throw new InvalidOperationException("Contour coverage intersections exceed the rasterization limit.");
                        }
                        retainedCrossings += crossings.Count;
                        scanlines.Add((high - low, rowCrossings.Count, crossings.Count));
                        rowCrossings.AddRange(crossings);
                    }
                }
                for (int tileLeft = left; tileLeft <= right; tileLeft += tileLength) {
                    int tileRight = Math.Min(right, tileLeft + tileLength - 1);
                    int count = tileRight - tileLeft + 1;
                    Array.Clear(coverage, 0, count);
                    foreach ((double weight, int start, int length) in scanlines) {
                        AccumulateContourIntervals(rowCrossings, start, length, fillRule, tileLeft, tileRight, weight, coverage);
                    }
                    for (int x = tileLeft; x <= tileRight; x++) {
                        double value = coverage[x - tileLeft];
                        if (value > 0D) BlendPixel(x, y, ApplyCoverage(paint(x + 0.5D, y + 0.5D), Math.Min(1D, value)));
                    }
                }
            }
        } finally {
#if NET8_0_OR_GREATER
            ArrayPool<double>.Shared.Return(coverage);
#endif
        }
    }

    private static void AddContourCrossings(IReadOnlyList<IReadOnlyList<OfficePoint>> contours, double y, List<ContourCrossing> crossings, bool secondShape = false) {
        // These interface-backed lists are visited for every sub-scanline. Indexed
        // traversal avoids boxing their enumerators for each glyph contour.
        for (int contourIndex = 0; contourIndex < contours.Count; contourIndex++) {
            IReadOnlyList<OfficePoint> contour = contours[contourIndex];
            if (contour.Count < 3) continue;
            OfficePoint start = contour[contour.Count - 1];
            for (int pointIndex = 0; pointIndex < contour.Count; pointIndex++) {
                OfficePoint end = contour[pointIndex];
                bool upward = start.Y <= y && end.Y > y;
                bool downward = start.Y > y && end.Y <= y;
                if (upward || downward) crossings.Add(new ContourCrossing(start.X + (y - start.Y) * (end.X - start.X) / (end.Y - start.Y), upward ? 1 : -1, secondShape));
                start = end;
            }
        }
    }

    private static void AccumulateContourIntervals(List<ContourCrossing> crossings, int startIndex, int count, OfficeFillRule rule, int left, int right, double weight, double[] coverage) {
        int winding = 0, secondWinding = 0, index = startIndex;
        int endIndex = startIndex + count;
        double previous = crossings[index].X;
        while (index < endIndex) {
            double x = crossings[index].X;
            bool inside = rule == OfficeFillRule.NonZero ? winding != 0 || secondWinding != 0 : (winding & 1) != 0 || (secondWinding & 1) != 0;
            if (inside && x > previous) {
                double start = Math.Max(left, previous), end = Math.Min(right + 1D, x);
                if (end > start) {
                    int first = (int)Math.Floor(start), last = (int)Math.Ceiling(end) - 1;
                    for (int pixel = first; pixel <= last; pixel++) {
                        coverage[pixel - left] += (Math.Min(end, pixel + 1D) - Math.Max(start, pixel)) * weight;
                    }
                }
            }
            do {
                int delta = rule == OfficeFillRule.NonZero ? crossings[index].WindingDelta : 1;
                if (crossings[index].SecondShape) secondWinding += delta;
                else winding += delta;
                index++;
            } while (index < endIndex && Math.Abs(crossings[index].X - x) <= 1E-9D);
            previous = x;
        }
    }

    private readonly struct ContourCrossing {
        internal ContourCrossing(double x, int windingDelta, bool secondShape) { X = x; WindingDelta = windingDelta; SecondShape = secondShape; }
        internal double X { get; }
        internal int WindingDelta { get; }
        internal bool SecondShape { get; }
    }

    private sealed class ContourCrossingComparer : IComparer<ContourCrossing> {
        internal static readonly ContourCrossingComparer Instance = new ContourCrossingComparer();
        public int Compare(ContourCrossing x, ContourCrossing y) => x.X.CompareTo(y.X);
    }
}

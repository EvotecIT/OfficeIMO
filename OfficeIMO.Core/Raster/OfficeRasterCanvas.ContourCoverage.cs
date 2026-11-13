#if NET8_0_OR_GREATER
using System.Buffers;
using System.Numerics;
#endif

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private const int ContourSubScanlines = 8;
    private const long MaximumContourCrossingWork = 512_000_000L;
    private const long MaximumContourRowCrossingWork = 16_000_000L;
    private const int MaximumRetainedContourCrossings = 1_000_000;

    private void FillContours(IReadOnlyList<IReadOnlyList<OfficePoint>> contours, OfficeColor color, OfficeFillRule fillRule) {
        if (color.A == 0) return;
        FillContourPaint(contours, fillRule, SolidPaint(color));
    }

    /// <summary>Returns a paint callback that the contour filler recognizes as one solid color.</summary>
    internal static Func<double, double, OfficeColor> SolidPaint(OfficeColor color) => new SolidContourPaint(color).Paint;

    private sealed class SolidContourPaint {
        internal SolidContourPaint(OfficeColor color) { Color = color; }
        internal OfficeColor Color { get; }
        internal OfficeColor Paint(double x, double y) => Color;
    }

    /// <summary>
    /// Accumulates horizontal area coverage before painting a pixel once. Vertex heights split
    /// the sub-scanlines, retaining thin horizontal details even between the regular sample rows.
    /// </summary>
    internal void FillContourPaint(IReadOnlyList<IReadOnlyList<OfficePoint>> contours, OfficeFillRule fillRule, Func<double, double, OfficeColor> paint,
        IReadOnlyList<IReadOnlyList<OfficePoint>>? unionContours = null) {
        if (contours == null || contours.Count == 0) return;
        double scaleX = 1D, scaleY = 1D;
        bool scaled = HasCoordinateScale;
        if (scaled) {
            scaleX = CoordinateScaleX;
            scaleY = CoordinateScaleY;
            if (paint.Target is not SolidContourPaint) {
                Func<double, double, OfficeColor> localPaint = paint;
                paint = (x, y) => localPaint(x / scaleX, y / scaleY);
            }
        }
        ContourCoverageWorkspace workspace = TakeContourCoverageWorkspace();
        try {
            FillContourPaint(contours, fillRule, paint, unionContours, workspace, null, scaled, scaleX, scaleY);
        } finally {
            ReturnContourCoverageWorkspace(workspace);
        }
    }

    // Edges are swept in the order of their lower end, so each sub-scanline visits only the
    // edges it crosses. Crossings keep the original arithmetic and are sorted by position,
    // so the accumulated coverage matches a scan over every edge of the pixel row.
    private void FillContourPaint(
        IReadOnlyList<IReadOnlyList<OfficePoint>> contours, OfficeFillRule fillRule, Func<double, double, OfficeColor> paint,
        IReadOnlyList<IReadOnlyList<OfficePoint>>? unionContours, ContourCoverageWorkspace workspace, IReadOnlyList<ClipContourSource>? intersectionClips = null,
        bool scaled = false, double scaleX = 1D, double scaleY = 1D) {
        double minX = double.PositiveInfinity, maxX = double.NegativeInfinity;
        long contourEdges = 0L;
        workspace.BoundaryCount = workspace.EdgeCount = 0;
        if (!CollectContourEdges(contours, 0, workspace, scaled, scaleX, scaleY, ref minX, ref maxX, ref contourEdges)) return;
        if (unionContours != null && !CollectContourEdges(unionContours, 1, workspace, scaled, scaleX, scaleY, ref minX, ref maxX, ref contourEdges)) return;
        if (intersectionClips != null) for (int clip = 1; clip < intersectionClips.Count; clip++) {
            if (!CollectContourEdges(intersectionClips[clip].Contours, clip, workspace, scaled, scaleX, scaleY, ref minX, ref maxX, ref contourEdges)) return;
        }
        if (workspace.BoundaryCount == 0 || maxX <= 0D || minX >= Width) return;
        double[] boundaries = workspace.Boundaries;
        int boundaryCount = SortDistinctContourBoundaries(boundaries, workspace.BoundaryCount);
        double minY = boundaries[0], maxY = boundaries[boundaryCount - 1];
        if (maxY <= 0D || minY >= Height) return;
        int left = (int)Math.Max(0D, Math.Floor(minX));
        int right = (int)Math.Min(Width - 1D, Math.Ceiling(maxX) - 1D);
        int top = (int)Math.Max(0D, Math.Floor(minY));
        int bottom = (int)Math.Min(Height - 1D, Math.Ceiling(maxY) - 1D);
        if (right < left || bottom < top) return;
        _transformedTextBudget.ChargePaintPixels((long)(right - left + 1) * (bottom - top + 1));
        int tileLength = Math.Min(right - left + 1, MaximumContourCoverageTileWidth);
        bool singleTile = right - left + 1 == tileLength;
        OfficeColor? solid = paint.Target is SolidContourPaint solidPaint ? solidPaint.Color : null;
#if NET8_0_OR_GREATER
        double[] coverage = ArrayPool<double>.Shared.Rent(tileLength);
#else
        double[] coverage = new double[tileLength];
#endif
        try {
            int edgeCount = workspace.EdgeCount;
            ContourEdge[] edges = workspace.Edges;
            double[] edgeLows = workspace.EdgeLows;
            Array.Sort(edgeLows, edges, 0, edgeCount);
            int[]? clipWindings = intersectionClips != null ? workspace.RentClipWindings(intersectionClips.Count) : null;
            int nextEdge = 0, activeCount = 0, boundaryIndex = 0;
            int scanlineCount = 0, rowCrossingCount = 0;
            long crossingWork = 0L;
            for (int y = top; y <= bottom; y++) {
                _cancellationToken.ThrowIfCancellationRequested();
                while (boundaryIndex < boundaryCount && boundaries[boundaryIndex] <= y) boundaryIndex++;
                int vertexStart = boundaryIndex;
                while (boundaryIndex < boundaryCount && boundaries[boundaryIndex] < y + 1D) boundaryIndex++;
                double[] rowBoundaries = workspace.RowBoundaries(ContourSubScanlines + 1 + boundaryIndex - vertexStart);
                int rowBoundaryCount = MergeRowBoundaries(y, boundaries, vertexStart, boundaryIndex, rowBoundaries);
                long rowWork = contourEdges * (rowBoundaryCount - 1L);
                if (rowWork > MaximumContourRowCrossingWork ||
                    rowWork > MaximumContourCrossingWork - crossingWork) {
                    throw new InvalidOperationException("Contour coverage work exceeds the rasterization limit.");
                }
                crossingWork += rowWork;
                _transformedTextBudget.ChargePaintCrossings(rowWork);
                scanlineCount = rowCrossingCount = 0;
                if (singleTile) Array.Clear(coverage, 0, tileLength);
                int retainedCrossings = 0;
                for (int index = 1; index < rowBoundaryCount; index++) {
                    double low = rowBoundaries[index - 1], high = rowBoundaries[index];
                    if (high <= low) continue;
                    double sample = (low + high) / 2D;
                    while (nextEdge < edgeCount && edgeLows[nextEdge] <= sample) {
                        ContourCrossing[] grown = workspace.ActiveCrossings(activeCount + 1);
                        grown[activeCount++] = new ContourCrossing(double.NegativeInfinity, nextEdge++);
                    }
                    activeCount = UpdateActiveCrossings(edges, workspace.Active, activeCount, sample);
                    ContourCrossing[] crossings = workspace.Active;
                    if (activeCount >= 2) {
                        if (activeCount > MaximumRetainedContourCrossings - retainedCrossings) {
                            throw new InvalidOperationException("Contour coverage intersections exceed the rasterization limit.");
                        }
                        retainedCrossings += activeCount;
                        // Ordinary glyphs fit one tile. Accumulate in the same
                        // scanline order without retaining every sample's crossings;
                        // wide shapes still need those samples for subsequent tiles.
                        if (singleTile) {
                            AccumulateContourIntervals(crossings, 0, activeCount, edges, fillRule,
                                left, right, high - low, coverage, intersectionClips, clipWindings);
                        } else {
                            workspace.AddScanline(ref scanlineCount, high - low, rowCrossingCount, activeCount);
                            ContourCrossing[] rowCrossings = workspace.RowCrossings(rowCrossingCount + activeCount);
                            Array.Copy(crossings, 0, rowCrossings, rowCrossingCount, activeCount);
                            rowCrossingCount += activeCount;
                        }
                    }
                }
                for (int tileLeft = left; tileLeft <= right; tileLeft += tileLength) {
                    int tileRight = Math.Min(right, tileLeft + tileLength - 1);
                    int count = tileRight - tileLeft + 1;
                    if (!singleTile) {
                        Array.Clear(coverage, 0, count);
                        for (int scanline = 0; scanline < scanlineCount; scanline++) {
                            (double weight, int start, int length) = workspace.Scanlines[scanline];
                            AccumulateContourIntervals(workspace.RowCrossingBuffer, start, length, edges, fillRule, tileLeft, tileRight, weight, coverage,
                                intersectionClips, clipWindings);
                        }
                    }
                    if (solid.HasValue) {
                        BlendSolidCoverage(y, tileLeft, tileRight, coverage, solid.Value);
                    } else {
                        for (int x = tileLeft; x <= tileRight; x++) {
                            double value = coverage[x - tileLeft];
                            if (value > 0D) BlendPixel(x, y, ApplyCoverage(paint(x + 0.5D, y + 0.5D), Math.Min(1D, value)));
                        }
                    }
                }
            }
        } finally {
#if NET8_0_OR_GREATER
            ArrayPool<double>.Shared.Return(coverage);
#endif
        }
    }

    /// <summary>Blends one row of solid-color coverage, checking cancellation once for the row.</summary>
    private void BlendSolidCoverage(int y, int tileLeft, int tileRight, double[] coverage, OfficeColor color) {
        _cancellationToken.ThrowIfCancellationRequested();
        OfficeRasterImage? image = _image;
        bool unclipped = _clipRegion == null;
        for (int x = tileLeft; x <= tileRight; x++) {
            double value = coverage[x - tileLeft];
            if (!(value > 0D)) continue;
            OfficeColor source = ApplyCoverage(color, Math.Min(1D, value));
            if (unclipped && image != null) {
                image.BlendPixel(x, y, source);
                continue;
            }
            double clip = GetClipCoverage(x, y);
            if (clip <= 0D) continue;
            if (clip < 1D) source = ApplyCoverage(source, clip);
            if (image != null) image.BlendPixel(x, y, source);
            else _target!.BlendPixel(x, y, source);
        }
    }

    // Retire edges that end at or below the sample, compute the remaining crossings, and keep
    // them ordered by position. Consecutive samples change the order only slightly.
    private static int UpdateActiveCrossings(ContourEdge[] edges, ContourCrossing[] active, int activeCount, double sample) {
        int count = 0;
        for (int index = 0; index < activeCount; index++) {
            int edgeIndex = active[index].Edge;
            ContourEdge edge = edges[edgeIndex];
            if (edge.High <= sample) continue;
            active[count++] = new ContourCrossing(edge.StartX + (sample - edge.StartY) * edge.DeltaX / edge.DeltaY, edgeIndex);
        }
        // Positions never contain NaN: every term is finite and the edge height is non-zero.
        int moves = 0, moveBudget = 8 * count + 64;
        for (int index = 1; index < count; index++) {
            ContourCrossing item = active[index];
            int target = index - 1;
            if (active[target].X <= item.X) continue;
            while (target >= 0 && active[target].X > item.X) {
                active[target + 1] = active[target];
                target--;
                moves++;
            }
            active[target + 1] = item;
            if (moves > moveBudget) {
                SortCrossings(active, count);
                break;
            }
        }
        return count;
    }

    private static void SortCrossings(ContourCrossing[] crossings, int count) {
#if NET8_0_OR_GREATER
        crossings.AsSpan(0, count).Sort(ContourCrossingComparer.Comparison);
#else
        Array.Sort(crossings, 0, count, ContourCrossingComparer.Instance);
#endif
    }

    private static int SortDistinctContourBoundaries(double[] boundaries, int length) {
        Array.Sort(boundaries, 0, length);
        int count = 0;
        for (int index = 0; index < length; index++) {
            double value = boundaries[index];
            if (count == 0 || value != boundaries[count - 1]) boundaries[count++] = value;
        }
        return count;
    }

    // The regular samples and the vertex heights inside the row are each ordered, so a merge
    // yields the same distinct sorted boundaries as sorting their concatenation.
    private static int MergeRowBoundaries(int y, double[] vertices, int vertexStart, int vertexEnd, double[] row) {
        int count = 0, sample = 0, vertex = vertexStart;
        double regular = y;
        while (sample <= ContourSubScanlines || vertex < vertexEnd) {
            double value;
            if (vertex >= vertexEnd || (sample <= ContourSubScanlines && regular <= vertices[vertex])) {
                value = regular;
                sample++;
                regular = y + sample / (double)ContourSubScanlines;
            } else {
                value = vertices[vertex++];
            }
            if (count == 0 || value != row[count - 1]) row[count++] = value;
        }
        return count;
    }

    private bool CollectContourEdges(IReadOnlyList<IReadOnlyList<OfficePoint>> contours, int group, ContourCoverageWorkspace workspace,
        bool scaled, double scaleX, double scaleY, ref double minX, ref double maxX, ref long contourEdges) {
        // Keep the shapes in source order without copying their contour lists
        // or boxing interface-backed point enumerators for every painted glyph.
        for (int contourIndex = 0; contourIndex < contours.Count; contourIndex++) {
            _cancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<OfficePoint> contour = contours[contourIndex];
            int pointCount = contour.Count;
            if (pointCount < 3) continue;
            contourEdges += pointCount;
            double[] boundaries = workspace.ReserveBoundaries(pointCount);
            int boundaryCount = workspace.BoundaryCount;
            OfficePoint first = default, start = default;
            for (int pointIndex = 0; pointIndex < pointCount; pointIndex++) {
                OfficePoint point = contour[pointIndex];
                if (scaled) point = new OfficePoint(point.X * scaleX, point.Y * scaleY);
                if (!IsFinite(point.X) || !IsFinite(point.Y)) return false;
                minX = Math.Min(minX, point.X);
                maxX = Math.Max(maxX, point.X);
                boundaries[boundaryCount++] = point.Y;
                if (pointIndex == 0) first = point;
                else workspace.AddEdge(start, point, group);
                start = point;
            }
            workspace.BoundaryCount = boundaryCount;
            workspace.AddEdge(start, first, group);
        }
        return true;
    }

    private struct ContourEdge {
        internal ContourEdge(OfficePoint start, OfficePoint end, int group) {
            StartX = start.X;
            StartY = start.Y;
            DeltaX = end.X - start.X;
            DeltaY = end.Y - start.Y;
            High = Math.Max(start.Y, end.Y);
            WindingDelta = start.Y < end.Y ? 1 : -1;
            Group = group;
        }
        internal double StartX;
        internal double StartY;
        internal double DeltaX;
        internal double DeltaY;
        internal double High;
        internal int WindingDelta;
        internal int Group;
    }

    private static void AccumulateContourIntervals(ContourCrossing[] crossings, int startIndex, int count, ContourEdge[] edges, OfficeFillRule rule,
        int left, int right, double weight, double[] coverage, IReadOnlyList<ClipContourSource>? intersectionClips, int[]? clipWindings) {
        if (intersectionClips != null) {
            AccumulateClipIntersection(crossings, startIndex, count, edges, intersectionClips, clipWindings!, left, right, weight, coverage);
            return;
        }
        int winding = 0, secondWinding = 0, index = startIndex;
        int endIndex = startIndex + count;
        double previous = crossings[index].X;
        while (index < endIndex) {
            double x = crossings[index].X;
            bool inside = rule == OfficeFillRule.NonZero ? winding != 0 || secondWinding != 0 : (winding & 1) != 0 || (secondWinding & 1) != 0;
            if (inside && x > previous) AccumulateInterval(Math.Max(left, previous), Math.Min(right + 1D, x), left, weight, coverage);
            do {
                ContourEdge edge = edges[crossings[index].Edge];
                int delta = rule == OfficeFillRule.NonZero ? edge.WindingDelta : 1;
                if (edge.Group == 1) secondWinding += delta;
                else winding += delta;
                index++;
            } while (index < endIndex && Math.Abs(crossings[index].X - x) <= ContourCrossingTolerance);
            previous = x;
        }
    }

    /// <summary>Adds the area of one horizontal interval, weighted by its sub-scanline height, to the covered pixels.</summary>
    private static void AccumulateInterval(double start, double end, int left, double weight, double[] coverage) {
        if (!(end > start)) return;
        int pixel = (int)Math.Floor(start), last = (int)Math.Ceiling(end) - 1;
        // Pixels wholly inside the interval add exactly (pixel + 1 - pixel) * weight, which is the weight itself.
        int innerFirst = (int)Math.Ceiling(start), innerLast = (int)Math.Floor(end) - 1;
        for (; pixel <= last && pixel < innerFirst; pixel++) {
            coverage[pixel - left] += (Math.Min(end, pixel + 1D) - Math.Max(start, pixel)) * weight;
        }
        int innerEnd = Math.Min(last, innerLast);
        if (pixel <= innerEnd) {
            AddWeight(coverage, pixel - left, innerEnd - pixel + 1, weight);
            pixel = innerEnd + 1;
        }
        for (; pixel <= last; pixel++) {
            coverage[pixel - left] += (Math.Min(end, pixel + 1D) - Math.Max(start, pixel)) * weight;
        }
    }

    private static void AddWeight(double[] coverage, int start, int length, double weight) {
        int index = start, end = start + length;
#if NET8_0_OR_GREATER
        if (Vector.IsHardwareAccelerated && length >= Vector<double>.Count) {
            var increment = new Vector<double>(weight);
            Span<double> span = coverage.AsSpan(start, length);
            int vectorEnd = length - Vector<double>.Count;
            int offset = 0;
            for (; offset <= vectorEnd; offset += Vector<double>.Count) {
                Span<double> slice = span.Slice(offset, Vector<double>.Count);
                (new Vector<double>(slice) + increment).CopyTo(slice);
            }
            index = start + offset;
        }
#endif
        for (; index < end; index++) coverage[index] += weight;
    }

    private readonly struct ContourCrossing {
        internal ContourCrossing(double x, int edge) { X = x; Edge = edge; }
        internal double X { get; }
        internal int Edge { get; }
    }

    private sealed class ContourCrossingComparer : IComparer<ContourCrossing> {
        internal static readonly ContourCrossingComparer Instance = new ContourCrossingComparer();
#if NET8_0_OR_GREATER
        internal static readonly Comparison<ContourCrossing> Comparison = Instance.Compare;
#endif
        public int Compare(ContourCrossing x, ContourCrossing y) => x.X.CompareTo(y.X);
    }
}

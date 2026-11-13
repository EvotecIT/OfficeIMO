#if NET8_0_OR_GREATER
using System.Buffers;
#endif

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    // Retain one small workspace per canvas. Complex paths can still use larger
    // transient buffers, but do not keep them alive with a long-lived canvas.
    private const int MaximumRetainedContourWorkspaceCapacity = 8192;
    private ContourCoverageWorkspace? _contourCoverageWorkspace;

    private ContourCoverageWorkspace TakeContourCoverageWorkspace() {
        ContourCoverageWorkspace? workspace = _contourCoverageWorkspace;
        _contourCoverageWorkspace = null;
        // Taking the workspace removes it from the cache, so nested paint cannot
        // overwrite scratch still in use by an outer contour operation.
        return workspace ?? new ContourCoverageWorkspace();
    }

    private void ReturnContourCoverageWorkspace(ContourCoverageWorkspace workspace) {
        workspace.Clear();
        _contourCoverageWorkspace ??= workspace;
    }

    private sealed class ContourCoverageWorkspace {
        private double[] _rowBoundaries = Array.Empty<double>();
        private ContourCrossing[] _rowCrossings = Array.Empty<ContourCrossing>();
        private (double Weight, int Start, int Count)[] _scanlines = Array.Empty<(double, int, int)>();
        private int[] _clipWindings = Array.Empty<int>();

        internal double[] Boundaries { get; private set; } = Array.Empty<double>();
        internal int BoundaryCount { get; set; }
        internal ContourEdge[] Edges { get; private set; } = Array.Empty<ContourEdge>();
        internal double[] EdgeLows { get; private set; } = Array.Empty<double>();
        internal int EdgeCount { get; set; }
        internal ContourCrossing[] Active { get; private set; } = Array.Empty<ContourCrossing>();
        internal ContourCrossing[] RowCrossingBuffer => _rowCrossings;
        internal (double Weight, int Start, int Count)[] Scanlines => _scanlines;

        internal double[] ReserveBoundaries(int additional) {
            if (Boundaries.Length - BoundaryCount < additional) Boundaries = Grow(Boundaries, BoundaryCount, checked(BoundaryCount + additional));
            return Boundaries;
        }

        /// <summary>Adds a contour edge; horizontal edges never cross a sub-scanline and are omitted.</summary>
        internal void AddEdge(OfficePoint start, OfficePoint end, int group) {
            if (start.Y == end.Y) return;
            if (EdgeCount == Edges.Length) {
                Edges = Grow(Edges, EdgeCount, EdgeCount + 1);
                EdgeLows = Grow(EdgeLows, EdgeCount, EdgeCount + 1);
            }
            Edges[EdgeCount] = new ContourEdge(start, end, group);
            EdgeLows[EdgeCount++] = Math.Min(start.Y, end.Y);
        }

        internal double[] RowBoundaries(int capacity) {
            if (_rowBoundaries.Length < capacity) _rowBoundaries = Grow(_rowBoundaries, 0, capacity);
            return _rowBoundaries;
        }

        internal ContourCrossing[] ActiveCrossings(int capacity) {
            if (Active.Length < capacity) Active = Grow(Active, Active.Length, capacity);
            return Active;
        }

        internal ContourCrossing[] RowCrossings(int capacity) {
            if (_rowCrossings.Length < capacity) _rowCrossings = Grow(_rowCrossings, _rowCrossings.Length, capacity);
            return _rowCrossings;
        }

        internal void AddScanline(ref int count, double weight, int start, int length) {
            if (count == _scanlines.Length) _scanlines = Grow(_scanlines, count, count + 1);
            _scanlines[count++] = (weight, start, length);
        }

        internal double[] Intervals { get; private set; } = Array.Empty<double>();

        internal double[] ReserveIntervals(int count) {
            if (Intervals.Length < count * 2) Intervals = Grow(Intervals, Intervals.Length, checked(count * 2));
            return Intervals;
        }

        internal int[] RentClipWindings(int count) {
            if (_clipWindings.Length < count) _clipWindings = new int[count];
            return _clipWindings;
        }

        // On .NET 8 and later every buffer comes from the shared pools and goes back after each
        // fill, so canvases and clip masks reuse the same scratch. Elsewhere a canvas retains
        // only small buffers.
        internal void Clear() {
            BoundaryCount = 0;
            EdgeCount = 0;
            Boundaries = Release(Boundaries);
            Edges = Release(Edges);
            EdgeLows = Release(EdgeLows);
            Active = Release(Active);
            _rowBoundaries = Release(_rowBoundaries);
            _rowCrossings = Release(_rowCrossings);
            _scanlines = Release(_scanlines);
            Intervals = Release(Intervals);
        }

        private static T[] Grow<T>(T[] source, int count, int required) {
            int capacity = Math.Max(Math.Max(required, 16), source.Length > int.MaxValue / 2 ? int.MaxValue : source.Length * 2);
#if NET8_0_OR_GREATER
            T[] replacement = ArrayPool<T>.Shared.Rent(capacity);
#else
            var replacement = new T[capacity];
#endif
            Array.Copy(source, replacement, count);
            Release(source);
            return replacement;
        }

        private static T[] Release<T>(T[] array) {
#if NET8_0_OR_GREATER
            if (array.Length > 0) ArrayPool<T>.Shared.Return(array);
            return Array.Empty<T>();
#else
            return array.Length > MaximumRetainedContourWorkspaceCapacity ? Array.Empty<T>() : array;
#endif
        }
    }
}

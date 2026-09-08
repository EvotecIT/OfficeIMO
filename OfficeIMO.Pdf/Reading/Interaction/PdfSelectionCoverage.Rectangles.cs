using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfSelectionCoverage {
    private static double CalculateRectangleUnionArea(IReadOnlyList<CoverageRectangle> rectangles, CancellationToken cancellationToken) {
        if (rectangles.Count == 0) return 0D;
        var yCoordinates = new List<double>(checked(rectangles.Count * 2));
        for (int index = 0; index < rectangles.Count; index++) {
            yCoordinates.Add(rectangles[index].Top);
            yCoordinates.Add(rectangles[index].Bottom);
        }
        yCoordinates.Sort();
        int uniqueCount = 0;
        for (int index = 0; index < yCoordinates.Count; index++) {
            if (uniqueCount == 0 || yCoordinates[index] != yCoordinates[uniqueCount - 1]) yCoordinates[uniqueCount++] = yCoordinates[index];
        }
        if (uniqueCount < yCoordinates.Count) yCoordinates.RemoveRange(uniqueCount, yCoordinates.Count - uniqueCount);
        var coordinateIndexes = new Dictionary<double, int>(yCoordinates.Count);
        for (int index = 0; index < yCoordinates.Count; index++) coordinateIndexes.Add(yCoordinates[index], index);
        var events = new List<CoverageEvent>(checked(rectangles.Count * 2));
        for (int index = 0; index < rectangles.Count; index++) {
            CoverageRectangle rectangle = rectangles[index];
            events.Add(new CoverageEvent(rectangle.Left, coordinateIndexes[rectangle.Top], coordinateIndexes[rectangle.Bottom] - 1, 1));
            events.Add(new CoverageEvent(rectangle.Right, coordinateIndexes[rectangle.Top], coordinateIndexes[rectangle.Bottom] - 1, -1));
        }
        events.Sort(static (first, second) => first.X.CompareTo(second.X));
        var coverage = new VerticalCoverageTree(yCoordinates);
        double area = 0D;
        double previousX = events[0].X;
        int eventIndex = 0;
        while (eventIndex < events.Count) {
            if ((eventIndex & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            double x = events[eventIndex].X;
            area += (x - previousX) * coverage.CoveredLength;
            while (eventIndex < events.Count && events[eventIndex].X == x) {
                CoverageEvent current = events[eventIndex++];
                coverage.Update(current.TopIndex, current.BottomIndex, current.Delta);
            }
            previousX = x;
        }
        return area;
    }

    private readonly struct CoverageEvent {
        internal CoverageEvent(double x, int topIndex, int bottomIndex, int delta) {
            X = x;
            TopIndex = topIndex;
            BottomIndex = bottomIndex;
            Delta = delta;
        }
        internal double X { get; }
        internal int TopIndex { get; }
        internal int BottomIndex { get; }
        internal int Delta { get; }
    }

    private sealed class VerticalCoverageTree {
        private readonly IReadOnlyList<double> _coordinates;
        private readonly int[] _coverageCounts;
        private readonly double[] _coveredLengths;

        internal VerticalCoverageTree(IReadOnlyList<double> coordinates) {
            _coordinates = coordinates;
            int intervalCount = coordinates.Count - 1;
            int storageSize = checked(Math.Max(1, intervalCount) * 4);
            _coverageCounts = new int[storageSize];
            _coveredLengths = new double[storageSize];
        }
        internal double CoveredLength => _coveredLengths[1];
        internal void Update(int firstInterval, int lastInterval, int delta) {
            if (firstInterval > lastInterval) return;
            Update(1, 0, _coordinates.Count - 2, firstInterval, lastInterval, delta);
        }
        private void Update(int node, int left, int right, int firstInterval, int lastInterval, int delta) {
            if (firstInterval <= left && right <= lastInterval) {
                _coverageCounts[node] += delta;
            } else {
                int middle = left + ((right - left) / 2);
                if (firstInterval <= middle) Update(node * 2, left, middle, firstInterval, lastInterval, delta);
                if (lastInterval > middle) Update((node * 2) + 1, middle + 1, right, firstInterval, lastInterval, delta);
            }
            if (_coverageCounts[node] > 0) {
                _coveredLengths[node] = _coordinates[right + 1] - _coordinates[left];
            } else if (left == right) {
                _coveredLengths[node] = 0D;
            } else {
                _coveredLengths[node] = _coveredLengths[node * 2] + _coveredLengths[(node * 2) + 1];
            }
        }
    }

    private readonly struct CoverageRectangle {
        internal CoverageRectangle(double left, double top, double right, double bottom) {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }
        internal double Left { get; }
        internal double Top { get; }
        internal double Right { get; }
        internal double Bottom { get; }
    }

}

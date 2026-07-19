using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private sealed class VisualContour {
        public VisualContour(List<OfficePoint> points, bool closed) {
            Points = points;
            Closed = closed;
            Bounds = VisualBounds.FromPoints(points);
        }

        public List<OfficePoint> Points { get; }
        public VisualBounds Bounds { get; }
        private bool Closed { get; }

        public int SegmentCount(bool closeForFill) =>
            Points.Count < 2
                ? 0
                : Points.Count - 1 + ((Closed || closeForFill) ? 1 : 0);

        public void GetSegment(
            int index,
            bool closeForFill,
            out OfficePoint start,
            out OfficePoint end) {
            start = Points[index];
            end = index + 1 < Points.Count
                ? Points[index + 1]
                : Points[0];
        }
    }

    private readonly struct VisualBounds {
        public VisualBounds(double left, double top, double right, double bottom) {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }

        public double Left { get; }
        public double Top { get; }
        public double Right { get; }
        public double Bottom { get; }
        public double Width => Right - Left;
        public double Height => Bottom - Top;
        public bool HasPositiveArea =>
            IsFinite(Left) &&
            IsFinite(Top) &&
            IsFinite(Right) &&
            IsFinite(Bottom) &&
            Width > VisualGeometryEpsilon &&
            Height > VisualGeometryEpsilon;

        public static VisualBounds FromContours(IReadOnlyList<VisualContour> contours) {
            if (contours.Count == 0) {
                return default;
            }

            VisualBounds result = contours[0].Bounds;
            for (int i = 1; i < contours.Count; i++) {
                VisualBounds bounds = contours[i].Bounds;
                result = new VisualBounds(
                    Math.Min(result.Left, bounds.Left),
                    Math.Min(result.Top, bounds.Top),
                    Math.Max(result.Right, bounds.Right),
                    Math.Max(result.Bottom, bounds.Bottom));
            }

            return result;
        }

        public static VisualBounds FromPoints(List<OfficePoint> points) {
            if (points.Count == 0) {
                return default;
            }

            double left = points[0].X;
            double top = points[0].Y;
            double right = left;
            double bottom = top;
            for (int i = 1; i < points.Count; i++) {
                OfficePoint point = points[i];
                left = Math.Min(left, point.X);
                top = Math.Min(top, point.Y);
                right = Math.Max(right, point.X);
                bottom = Math.Max(bottom, point.Y);
            }

            return new VisualBounds(left, top, right, bottom);
        }

        public static VisualBounds FromSegment(OfficePoint start, OfficePoint end) =>
            new VisualBounds(
                Math.Min(start.X, end.X),
                Math.Min(start.Y, end.Y),
                Math.Max(start.X, end.X),
                Math.Max(start.Y, end.Y));

        public VisualBounds Expand(double amount) =>
            new VisualBounds(
                Left - amount,
                Top - amount,
                Right + amount,
                Bottom + amount);

        public bool TryIntersectPositive(
            VisualBounds other,
            out VisualBounds intersection) {
            intersection = new VisualBounds(
                Math.Max(Left, other.Left),
                Math.Max(Top, other.Top),
                Math.Min(Right, other.Right),
                Math.Min(Bottom, other.Bottom));
            return intersection.HasPositiveArea;
        }

        public bool IntersectsInclusive(VisualBounds other) =>
            Left <= other.Right + VisualGeometryEpsilon &&
            Right >= other.Left - VisualGeometryEpsilon &&
            Top <= other.Bottom + VisualGeometryEpsilon &&
            Bottom >= other.Top - VisualGeometryEpsilon;

        public bool ContainsInclusive(OfficePoint point) =>
            point.X >= Left - VisualGeometryEpsilon &&
            point.X <= Right + VisualGeometryEpsilon &&
            point.Y >= Top - VisualGeometryEpsilon &&
            point.Y <= Bottom + VisualGeometryEpsilon;

        public bool ContainsStrict(OfficePoint point) =>
            point.X > Left + VisualGeometryEpsilon &&
            point.X < Right - VisualGeometryEpsilon &&
            point.Y > Top + VisualGeometryEpsilon &&
            point.Y < Bottom - VisualGeometryEpsilon;
    }

    private static bool TryGetUnitNormal(
        OfficePoint start,
        OfficePoint end,
        double distance,
        out double normalX,
        out double normalY) {
        double deltaX = end.X - start.X;
        double deltaY = end.Y - start.Y;
        double length = Math.Sqrt((deltaX * deltaX) + (deltaY * deltaY));
        if (!IsFinite(length) || length <= VisualGeometryEpsilon) {
            normalX = 0D;
            normalY = 0D;
            return false;
        }

        normalX = -deltaY / length * distance;
        normalY = deltaX / length * distance;
        return true;
    }

    private static bool TryGetSegmentIntersection(
        OfficePoint firstStart,
        OfficePoint firstEnd,
        OfficePoint secondStart,
        OfficePoint secondEnd,
        out OfficePoint intersection) {
        double firstX = firstEnd.X - firstStart.X;
        double firstY = firstEnd.Y - firstStart.Y;
        double secondX = secondEnd.X - secondStart.X;
        double secondY = secondEnd.Y - secondStart.Y;
        double denominator = (firstX * secondY) - (firstY * secondX);
        if (Math.Abs(denominator) <= VisualGeometryEpsilon) {
            intersection = default;
            return false;
        }

        double offsetX = secondStart.X - firstStart.X;
        double offsetY = secondStart.Y - firstStart.Y;
        double firstAmount = ((offsetX * secondY) - (offsetY * secondX)) / denominator;
        double secondAmount = ((offsetX * firstY) - (offsetY * firstX)) / denominator;
        if (firstAmount < -VisualGeometryEpsilon ||
            firstAmount > 1D + VisualGeometryEpsilon ||
            secondAmount < -VisualGeometryEpsilon ||
            secondAmount > 1D + VisualGeometryEpsilon) {
            intersection = default;
            return false;
        }

        intersection = new OfficePoint(
            firstStart.X + (firstAmount * firstX),
            firstStart.Y + (firstAmount * firstY));
        return true;
    }

    private static bool PointOnSegment(
        OfficePoint point,
        OfficePoint start,
        OfficePoint end) {
        if (Math.Abs(Cross(start, end, point)) > VisualGeometryEpsilon) {
            return false;
        }

        return point.X >= Math.Min(start.X, end.X) - VisualGeometryEpsilon &&
            point.X <= Math.Max(start.X, end.X) + VisualGeometryEpsilon &&
            point.Y >= Math.Min(start.Y, end.Y) - VisualGeometryEpsilon &&
            point.Y <= Math.Max(start.Y, end.Y) + VisualGeometryEpsilon;
    }

    private static double SegmentDistanceSquared(
        OfficePoint firstStart,
        OfficePoint firstEnd,
        OfficePoint secondStart,
        OfficePoint secondEnd) {
        if (SegmentsIntersect(firstStart, firstEnd, secondStart, secondEnd)) {
            return 0D;
        }

        return Math.Min(
            Math.Min(
                PointSegmentDistanceSquared(firstStart, secondStart, secondEnd),
                PointSegmentDistanceSquared(firstEnd, secondStart, secondEnd)),
            Math.Min(
                PointSegmentDistanceSquared(secondStart, firstStart, firstEnd),
                PointSegmentDistanceSquared(secondEnd, firstStart, firstEnd)));
    }

    private static bool SegmentsIntersect(
        OfficePoint firstStart,
        OfficePoint firstEnd,
        OfficePoint secondStart,
        OfficePoint secondEnd) {
        double firstCrossStart = Cross(firstStart, firstEnd, secondStart);
        double firstCrossEnd = Cross(firstStart, firstEnd, secondEnd);
        double secondCrossStart = Cross(secondStart, secondEnd, firstStart);
        double secondCrossEnd = Cross(secondStart, secondEnd, firstEnd);
        if (((firstCrossStart > VisualGeometryEpsilon && firstCrossEnd < -VisualGeometryEpsilon) ||
             (firstCrossStart < -VisualGeometryEpsilon && firstCrossEnd > VisualGeometryEpsilon)) &&
            ((secondCrossStart > VisualGeometryEpsilon && secondCrossEnd < -VisualGeometryEpsilon) ||
             (secondCrossStart < -VisualGeometryEpsilon && secondCrossEnd > VisualGeometryEpsilon))) {
            return true;
        }

        return (Math.Abs(firstCrossStart) <= VisualGeometryEpsilon &&
                PointOnSegment(secondStart, firstStart, firstEnd)) ||
            (Math.Abs(firstCrossEnd) <= VisualGeometryEpsilon &&
             PointOnSegment(secondEnd, firstStart, firstEnd)) ||
            (Math.Abs(secondCrossStart) <= VisualGeometryEpsilon &&
             PointOnSegment(firstStart, secondStart, secondEnd)) ||
            (Math.Abs(secondCrossEnd) <= VisualGeometryEpsilon &&
             PointOnSegment(firstEnd, secondStart, secondEnd));
    }

    private static double PointSegmentDistanceSquared(
        OfficePoint point,
        OfficePoint start,
        OfficePoint end) {
        double deltaX = end.X - start.X;
        double deltaY = end.Y - start.Y;
        double lengthSquared = (deltaX * deltaX) + (deltaY * deltaY);
        if (lengthSquared <= VisualGeometryEpsilon) {
            double pointDeltaX = point.X - start.X;
            double pointDeltaY = point.Y - start.Y;
            return (pointDeltaX * pointDeltaX) + (pointDeltaY * pointDeltaY);
        }

        double projection = (((point.X - start.X) * deltaX) +
            ((point.Y - start.Y) * deltaY)) /
            lengthSquared;
        projection = Math.Max(0D, Math.Min(1D, projection));
        double closestX = start.X + (projection * deltaX);
        double closestY = start.Y + (projection * deltaY);
        double distanceX = point.X - closestX;
        double distanceY = point.Y - closestY;
        return (distanceX * distanceX) + (distanceY * distanceY);
    }

    private static double Cross(
        OfficePoint start,
        OfficePoint end,
        OfficePoint point) =>
        ((end.X - start.X) * (point.Y - start.Y)) -
        ((end.Y - start.Y) * (point.X - start.X));

    private static bool PointsEqual(OfficePoint left, OfficePoint right) =>
        Math.Abs(left.X - right.X) <= VisualGeometryEpsilon &&
        Math.Abs(left.Y - right.Y) <= VisualGeometryEpsilon;
}

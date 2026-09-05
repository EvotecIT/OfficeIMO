namespace OfficeIMO.Pdf;

/// <summary>Exact destructive geometry carried by a bounded redaction area.</summary>
internal sealed class PdfRedactionGeometry {
    private const double Tolerance = 0.000001D;

    private PdfRedactionGeometry(PdfRedactionRegionKind kind, PdfRedactionPoint[] points, double strokeWidth) {
        Kind = kind;
        Points = points;
        StrokeWidth = strokeWidth;
    }

    internal PdfRedactionRegionKind Kind { get; }
    internal IReadOnlyList<PdfRedactionPoint> Points { get; }
    internal double StrokeWidth { get; }

    internal static PdfRedactionGeometry Polygon(PdfRedactionRegionKind kind, PdfRedactionPoint[] points) =>
        new PdfRedactionGeometry(kind, (PdfRedactionPoint[])points.Clone(), 0D);

    internal static PdfRedactionGeometry FreehandSegment(PdfRedactionPoint start, PdfRedactionPoint end, double strokeWidth) =>
        new PdfRedactionGeometry(PdfRedactionRegionKind.Freehand, new[] { start, end }, strokeWidth);

    internal static bool IsSimplePolygon(IReadOnlyList<PdfRedactionPoint> points) {
        if (points.Count < 3) return false;
        for (int index = 0; index < points.Count; index++) {
            PdfRedactionPoint first = points[index];
            PdfRedactionPoint second = points[(index + 1) % points.Count];
            if (DistanceSquared(first, second) <= Square(Tolerance)) return false;
            for (int other = index + 1; other < points.Count; other++) {
                int nextOther = (other + 1) % points.Count;
                if (other == index || other == (index + 1) % points.Count || nextOther == index) continue;
                if (SegmentsIntersect(first, second, points[other], points[nextOther])) return false;
            }
        }
        return Math.Abs(SignedArea(points)) > Tolerance;
    }

    internal bool IntersectsRectangle(double x, double y, double width, double height) {
        if (width <= 0D || height <= 0D) return false;
        return Kind == PdfRedactionRegionKind.Freehand
            ? DistanceSquaredSegmentToRectangle(Points[0], Points[1], x, y, width, height) <= Square(StrokeWidth / 2D) + Tolerance
            : PolygonIntersectsRectangle(Points, x, y, width, height);
    }

    internal bool ContainsRectangle(double x, double y, double width, double height) {
        if (width <= 0D || height <= 0D) return false;
        PdfRedactionPoint[] corners = RectangleCorners(x, y, width, height);
        if (Kind == PdfRedactionRegionKind.Freehand) {
            double radiusSquared = Square(StrokeWidth / 2D);
            return corners.All(point => DistanceSquaredToSegment(point, Points[0], Points[1]) <= radiusSquared + Tolerance);
        }

        if (!corners.All(point => ContainsPoint(point.X, point.Y))) return false;
        // A concave polygon can contain all corners while a notch crosses the rectangle.
        for (int index = 0; index < Points.Count; index++) {
            PdfRedactionPoint first = Points[index];
            PdfRedactionPoint second = Points[(index + 1) % Points.Count];
            if (SegmentCrossesRectangleBoundary(first, second, x, y, width, height)) return false;
        }
        return true;
    }

    internal bool ContainsPoint(double x, double y) {
        if (Kind == PdfRedactionRegionKind.Freehand) {
            return DistanceSquaredToSegment(new PdfRedactionPoint(x, y), Points[0], Points[1]) <= Square(StrokeWidth / 2D) + Tolerance;
        }

        bool inside = false;
        for (int current = 0, previous = Points.Count - 1; current < Points.Count; previous = current++) {
            PdfRedactionPoint a = Points[previous];
            PdfRedactionPoint b = Points[current];
            if (PointOnSegment(x, y, a, b)) return true;
            if ((a.Y > y) != (b.Y > y) && x < (b.X - a.X) * (y - a.Y) / (b.Y - a.Y) + a.X) inside = !inside;
        }
        return inside;
    }

    internal bool IntersectsQuadrilateral(
        PdfRedactionPoint first,
        PdfRedactionPoint second,
        PdfRedactionPoint third,
        PdfRedactionPoint fourth) {
        if (Kind == PdfRedactionRegionKind.Freehand) {
            double radiusSquared = Square(StrokeWidth / 2D);
            return PointInConvexQuadrilateral(Points[0], first, second, third, fourth) ||
                PointInConvexQuadrilateral(Points[1], first, second, third, fourth) ||
                DistanceSquaredToSegment(first, Points[0], Points[1]) <= radiusSquared + Tolerance ||
                DistanceSquaredToSegment(second, Points[0], Points[1]) <= radiusSquared + Tolerance ||
                DistanceSquaredToSegment(third, Points[0], Points[1]) <= radiusSquared + Tolerance ||
                DistanceSquaredToSegment(fourth, Points[0], Points[1]) <= radiusSquared + Tolerance ||
                DistanceSquaredBetweenSegments(Points[0], Points[1], first, second) <= radiusSquared + Tolerance ||
                DistanceSquaredBetweenSegments(Points[0], Points[1], second, third) <= radiusSquared + Tolerance ||
                DistanceSquaredBetweenSegments(Points[0], Points[1], third, fourth) <= radiusSquared + Tolerance ||
                DistanceSquaredBetweenSegments(Points[0], Points[1], fourth, first) <= radiusSquared + Tolerance;
        }

        if (ContainsPoint(first.X, first.Y) || ContainsPoint(second.X, second.Y) ||
            ContainsPoint(third.X, third.Y) || ContainsPoint(fourth.X, fourth.Y)) return true;
        for (int index = 0; index < Points.Count; index++) {
            PdfRedactionPoint point = Points[index];
            if (PointInConvexQuadrilateral(point, first, second, third, fourth)) return true;
            PdfRedactionPoint next = Points[(index + 1) % Points.Count];
            if (SegmentsIntersect(point, next, first, second) ||
                SegmentsIntersect(point, next, second, third) ||
                SegmentsIntersect(point, next, third, fourth) ||
                SegmentsIntersect(point, next, fourth, first)) return true;
        }
        return false;
    }

    internal static bool RectangleIntersectsQuadrilateral(
        double x,
        double y,
        double width,
        double height,
        PdfRedactionPoint first,
        PdfRedactionPoint second,
        PdfRedactionPoint third,
        PdfRedactionPoint fourth) {
        double right = x + width;
        double top = y + height;
        if (PointInsideRectangle(first) || PointInsideRectangle(second) ||
            PointInsideRectangle(third) || PointInsideRectangle(fourth)) return true;

        var bottomLeft = new PdfRedactionPoint(x, y);
        var bottomRight = new PdfRedactionPoint(right, y);
        var topRight = new PdfRedactionPoint(right, top);
        var topLeft = new PdfRedactionPoint(x, top);
        if (PointInConvexQuadrilateral(bottomLeft, first, second, third, fourth) ||
            PointInConvexQuadrilateral(bottomRight, first, second, third, fourth) ||
            PointInConvexQuadrilateral(topRight, first, second, third, fourth) ||
            PointInConvexQuadrilateral(topLeft, first, second, third, fourth)) return true;

        return RectangleEdgeIntersectsQuadrilateral(bottomLeft, bottomRight) ||
            RectangleEdgeIntersectsQuadrilateral(bottomRight, topRight) ||
            RectangleEdgeIntersectsQuadrilateral(topRight, topLeft) ||
            RectangleEdgeIntersectsQuadrilateral(topLeft, bottomLeft);

        bool PointInsideRectangle(PdfRedactionPoint point) =>
            point.X >= x - Tolerance && point.X <= right + Tolerance &&
            point.Y >= y - Tolerance && point.Y <= top + Tolerance;

        bool RectangleEdgeIntersectsQuadrilateral(PdfRedactionPoint start, PdfRedactionPoint end) =>
            SegmentsIntersect(start, end, first, second) ||
            SegmentsIntersect(start, end, second, third) ||
            SegmentsIntersect(start, end, third, fourth) ||
            SegmentsIntersect(start, end, fourth, first);
    }

    private bool PolygonIntersectsRectangle(IReadOnlyList<PdfRedactionPoint> polygon, double x, double y, double width, double height) {
        double right = x + width;
        double top = y + height;
        if (polygon.Any(point => point.X >= x && point.X <= right && point.Y >= y && point.Y <= top)) return true;
        PdfRedactionPoint[] corners = RectangleCorners(x, y, width, height);
        if (corners.Any(point => ContainsPoint(point.X, point.Y))) return true;
        for (int index = 0; index < polygon.Count; index++) {
            PdfRedactionPoint first = polygon[index];
            PdfRedactionPoint second = polygon[(index + 1) % polygon.Count];
            for (int edge = 0; edge < corners.Length; edge++) {
                if (SegmentsIntersect(first, second, corners[edge], corners[(edge + 1) % corners.Length])) return true;
            }
        }
        return false;
    }

    private static bool SegmentCrossesRectangleBoundary(PdfRedactionPoint a, PdfRedactionPoint b, double x, double y, double width, double height) {
        PdfRedactionPoint[] corners = RectangleCorners(x, y, width, height);
        for (int edge = 0; edge < corners.Length; edge++) {
            if (SegmentsIntersect(a, b, corners[edge], corners[(edge + 1) % corners.Length]) &&
                !PointOnSegment(corners[edge].X, corners[edge].Y, a, b) &&
                !PointOnSegment(corners[(edge + 1) % corners.Length].X, corners[(edge + 1) % corners.Length].Y, a, b)) return true;
        }
        return false;
    }

    private static double DistanceSquaredSegmentToRectangle(PdfRedactionPoint a, PdfRedactionPoint b, double x, double y, double width, double height) {
        if (a.X >= x && a.X <= x + width && a.Y >= y && a.Y <= y + height ||
            b.X >= x && b.X <= x + width && b.Y >= y && b.Y <= y + height) return 0D;
        PdfRedactionPoint[] corners = RectangleCorners(x, y, width, height);
        double minimum = double.PositiveInfinity;
        for (int edge = 0; edge < corners.Length; edge++) {
            minimum = Math.Min(minimum, DistanceSquaredBetweenSegments(a, b, corners[edge], corners[(edge + 1) % corners.Length]));
        }
        return minimum;
    }

    private static double DistanceSquaredBetweenSegments(PdfRedactionPoint a, PdfRedactionPoint b, PdfRedactionPoint c, PdfRedactionPoint d) {
        if (SegmentsIntersect(a, b, c, d)) return 0D;
        return Math.Min(
            Math.Min(DistanceSquaredToSegment(a, c, d), DistanceSquaredToSegment(b, c, d)),
            Math.Min(DistanceSquaredToSegment(c, a, b), DistanceSquaredToSegment(d, a, b)));
    }

    private static double DistanceSquaredToSegment(PdfRedactionPoint point, PdfRedactionPoint start, PdfRedactionPoint end) {
        double dx = end.X - start.X;
        double dy = end.Y - start.Y;
        double lengthSquared = dx * dx + dy * dy;
        if (lengthSquared <= Tolerance) return Square(point.X - start.X) + Square(point.Y - start.Y);
        double projection = ((point.X - start.X) * dx + (point.Y - start.Y) * dy) / lengthSquared;
        projection = Math.Max(0D, Math.Min(1D, projection));
        double closestX = start.X + projection * dx;
        double closestY = start.Y + projection * dy;
        return Square(point.X - closestX) + Square(point.Y - closestY);
    }

    private static bool SegmentsIntersect(PdfRedactionPoint a, PdfRedactionPoint b, PdfRedactionPoint c, PdfRedactionPoint d) {
        double o1 = Orientation(a, b, c);
        double o2 = Orientation(a, b, d);
        double o3 = Orientation(c, d, a);
        double o4 = Orientation(c, d, b);
        if (Math.Abs(o1) <= Tolerance && PointOnSegment(c.X, c.Y, a, b) ||
            Math.Abs(o2) <= Tolerance && PointOnSegment(d.X, d.Y, a, b) ||
            Math.Abs(o3) <= Tolerance && PointOnSegment(a.X, a.Y, c, d) ||
            Math.Abs(o4) <= Tolerance && PointOnSegment(b.X, b.Y, c, d)) return true;
        return (o1 > 0D) != (o2 > 0D) && (o3 > 0D) != (o4 > 0D);
    }

    private static bool PointInConvexQuadrilateral(
        PdfRedactionPoint point,
        PdfRedactionPoint first,
        PdfRedactionPoint second,
        PdfRedactionPoint third,
        PdfRedactionPoint fourth) {
        double firstOrientation = Orientation(first, second, point);
        double secondOrientation = Orientation(second, third, point);
        double thirdOrientation = Orientation(third, fourth, point);
        double fourthOrientation = Orientation(fourth, first, point);
        bool hasNegative = firstOrientation < -Tolerance || secondOrientation < -Tolerance || thirdOrientation < -Tolerance || fourthOrientation < -Tolerance;
        bool hasPositive = firstOrientation > Tolerance || secondOrientation > Tolerance || thirdOrientation > Tolerance || fourthOrientation > Tolerance;
        return !(hasNegative && hasPositive);
    }

    private static double Orientation(PdfRedactionPoint a, PdfRedactionPoint b, PdfRedactionPoint c) =>
        (b.X - a.X) * (c.Y - a.Y) - (b.Y - a.Y) * (c.X - a.X);

    private static double SignedArea(IReadOnlyList<PdfRedactionPoint> points) {
        double twiceArea = 0D;
        for (int index = 0; index < points.Count; index++) {
            PdfRedactionPoint current = points[index];
            PdfRedactionPoint next = points[(index + 1) % points.Count];
            twiceArea += current.X * next.Y - next.X * current.Y;
        }
        return twiceArea / 2D;
    }

    private static double DistanceSquared(PdfRedactionPoint left, PdfRedactionPoint right) =>
        Square(left.X - right.X) + Square(left.Y - right.Y);

    private static bool PointOnSegment(double x, double y, PdfRedactionPoint a, PdfRedactionPoint b) =>
        Math.Abs(Orientation(a, b, new PdfRedactionPoint(x, y))) <= Tolerance &&
        x >= Math.Min(a.X, b.X) - Tolerance && x <= Math.Max(a.X, b.X) + Tolerance &&
        y >= Math.Min(a.Y, b.Y) - Tolerance && y <= Math.Max(a.Y, b.Y) + Tolerance;

    private static PdfRedactionPoint[] RectangleCorners(double x, double y, double width, double height) => new[] {
        new PdfRedactionPoint(x, y), new PdfRedactionPoint(x + width, y),
        new PdfRedactionPoint(x + width, y + height), new PdfRedactionPoint(x, y + height)
    };

    private static double Square(double value) => value * value;
}

using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Measures union coverage of visual selection quads without counting duplicate spans twice.</summary>
internal static partial class PdfSelectionCoverage {
    internal static bool Covers(PdfSelectionQuad target, IReadOnlyList<PdfSelectionQuad> regions, double threshold,
        Action<long> consumeWork, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var targetPoints = Points(target);
        double targetArea = Area(targetPoints);
        if (targetArea <= 0D || regions.Count == 0) return false;
        double required = targetArea * threshold;
        bool targetIsRectangle = IsRectangle(target);
        var rectangles = new List<CoverageRectangle>();
        var polygons = new List<List<OfficePoint>>();
        double summedArea = 0D;
        for (int index = 0; index < regions.Count; index++) {
            consumeWork(1);
            cancellationToken.ThrowIfCancellationRequested();
            PdfSelectionQuad region = regions[index];
            double left = Math.Max(target.Left, region.Left), top = Math.Max(target.Top, region.Top);
            double right = Math.Min(target.Right, region.Right), bottom = Math.Min(target.Bottom, region.Bottom);
            if (right <= left || bottom <= top) continue;
            double area;
            if (targetIsRectangle && IsRectangle(region)) {
                area = (right - left) * (bottom - top);
                rectangles.Add(new CoverageRectangle(left, top, right, bottom));
            } else {
                // Convex quad clipping has at most eight output vertices and four clip edges.
                consumeWork(64);
                List<OfficePoint> clipped = PdfPageClipPath.ClipPolygonToConvexPolygon(Points(region), targetPoints, null);
                area = Area(clipped);
                if (area <= 0D) continue;
                polygons.Add(clipped);
            }
            if (area >= required) return true;
            summedArea += area;
        }
        if (summedArea < required || summedArea <= 0D) return false;
        if (polygons.Count == 0) return CalculateRectangleUnionArea(rectangles, cancellationToken) >= required;
        foreach (CoverageRectangle rectangle in rectangles) polygons.Add(new List<OfficePoint> {
            new OfficePoint(rectangle.Left, rectangle.Top), new OfficePoint(rectangle.Right, rectangle.Top),
            new OfficePoint(rectangle.Right, rectangle.Bottom), new OfficePoint(rectangle.Left, rectangle.Bottom)
        });
        return PolygonUnionArea(polygons, required, consumeWork, cancellationToken) >= required;
    }

    private static double PolygonUnionArea(List<List<OfficePoint>> polygons, double stopAt,
        Action<long> consumeWork, CancellationToken token) {
        var edges = new List<(OfficePoint Start, OfficePoint End)>();
        var xs = new List<double>();
        foreach (List<OfficePoint> polygon in polygons) {
            consumeWork(polygon.Count);
            for (int i = 0; i < polygon.Count; i++) {
                xs.Add(polygon[i].X);
                edges.Add((polygon[i], polygon[(i + 1) % polygon.Count]));
            }
        }
        // Edge crossings partition X into slabs on which the union's vertical length is linear.
        // Integrating at the slab midpoint is therefore exact for these straight-edge polygons.
        for (int i = 0; i < edges.Count; i++) {
            token.ThrowIfCancellationRequested();
            for (int j = i + 1; j < edges.Count; j++) {
                consumeWork(1);
                var a = edges[i]; var b = edges[j];
                double ax = a.End.X - a.Start.X, ay = a.End.Y - a.Start.Y;
                double bx = b.End.X - b.Start.X, by = b.End.Y - b.Start.Y;
                double denominator = ax * by - ay * bx;
                if (denominator == 0D) continue;
                double dx = b.Start.X - a.Start.X, dy = b.Start.Y - a.Start.Y;
                double t = (dx * by - dy * bx) / denominator, u = (dx * ay - dy * ax) / denominator;
                if (t > 0D && t < 1D && u > 0D && u < 1D) xs.Add(a.Start.X + t * ax);
            }
        }
        consumeWork((long)xs.Count * Math.Max(1, (int)Math.Ceiling(Math.Log(Math.Max(2, xs.Count), 2D))));
        xs.Sort();
        double area = 0D;
        var intervals = new List<(double Top, double Bottom)>(polygons.Count);
        for (int i = 1; i < xs.Count; i++) {
            token.ThrowIfCancellationRequested();
            double span = xs[i] - xs[i - 1];
            if (span <= 0D) continue;
            double x = xs[i - 1] + span / 2D;
            intervals.Clear();
            foreach (List<OfficePoint> polygon in polygons) {
                consumeWork(polygon.Count);
                double top = double.PositiveInfinity, bottom = double.NegativeInfinity;
                for (int p = 0; p < polygon.Count; p++) {
                    OfficePoint first = polygon[p], second = polygon[(p + 1) % polygon.Count];
                    if (x <= Math.Min(first.X, second.X) || x >= Math.Max(first.X, second.X)) continue;
                    double y = first.Y + (x - first.X) * (second.Y - first.Y) / (second.X - first.X);
                    top = Math.Min(top, y); bottom = Math.Max(bottom, y);
                }
                if (bottom > top) intervals.Add((top, bottom));
            }
            consumeWork((long)intervals.Count * Math.Max(1, (int)Math.Ceiling(Math.Log(Math.Max(2, intervals.Count), 2D))));
            intervals.Sort(static (first, second) => first.Top.CompareTo(second.Top));
            double covered = 0D, end = double.NegativeInfinity;
            foreach (var interval in intervals) {
                covered += Math.Max(0D, interval.Bottom - Math.Max(end, interval.Top));
                end = Math.Max(end, interval.Bottom);
            }
            area += span * covered;
            if (area >= stopAt) return area;
        }
        return area;
    }

    private static bool IsRectangle(PdfSelectionQuad quad) =>
        quad.TopLeft.X == quad.BottomLeft.X && quad.TopRight.X == quad.BottomRight.X &&
        quad.TopLeft.Y == quad.TopRight.Y && quad.BottomLeft.Y == quad.BottomRight.Y ||
        quad.TopLeft.X == quad.TopRight.X && quad.BottomLeft.X == quad.BottomRight.X &&
        quad.TopLeft.Y == quad.BottomLeft.Y && quad.TopRight.Y == quad.BottomRight.Y;

    private static List<OfficePoint> Points(PdfSelectionQuad quad) => new List<OfficePoint> {
        new OfficePoint(quad.TopLeft.X, quad.TopLeft.Y), new OfficePoint(quad.TopRight.X, quad.TopRight.Y),
        new OfficePoint(quad.BottomRight.X, quad.BottomRight.Y), new OfficePoint(quad.BottomLeft.X, quad.BottomLeft.Y)
    };

    private static double Area(List<OfficePoint> points) {
        double twiceArea = 0D;
        for (int i = 0; i < points.Count; i++) {
            OfficePoint next = points[(i + 1) % points.Count];
            twiceArea += points[i].X * next.Y - next.X * points[i].Y;
        }
        return Math.Abs(twiceArea) / 2D;
    }
}

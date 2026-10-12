using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

internal readonly partial struct OfficeTextInkClip {
    private OfficeTextInkClip(IReadOnlyList<OfficePoint> polygon, double orientation) {
        this = default; _polygon = polygon; _orientation = orientation;
    }

    private OfficeTextInkClip(IReadOnlyList<IReadOnlyList<OfficePoint>> contours, OfficeFillRule fillRule) {
        this = default; FilledContours = contours; FillRule = fillRule;
    }

    internal static OfficeTextInkClip FromFilledContours(IReadOnlyList<IReadOnlyList<OfficePoint>> contours, OfficeFillRule fillRule) =>
        new OfficeTextInkClip(contours, fillRule);

    internal static bool TryCreatePath(OfficeClipPath path, OfficeTransform transform,
        CancellationToken cancellationToken, out OfficeTextInkClip clip) {
        if (TryCreateConvexPath(path, transform, cancellationToken, out clip)) return true;
        if (path.Commands.Count > 512) return false;
        bool valid = true;
        var contours = OfficeClipPathGeometry.CreateContours(path, points => {
            var mapped = new List<OfficePoint>(points.Count);
            foreach (OfficePoint point in points) {
                cancellationToken.ThrowIfCancellationRequested();
                OfficePoint value = transform.TransformPoint(point);
                valid &= Finite(value.X) && Finite(value.Y);
                mapped.Add(value);
            }
            return mapped;
        });
        int vertices = 0;
        foreach (var contour in contours) { vertices += contour.Count; if (vertices > 512) return false; }
        if (!valid || (contours.Count == 0 && path.Kind != OfficeClipPathKind.Empty)) return false;
        clip = new OfficeTextInkClip(contours, path.FillRule);
        return true;
    }

    internal static bool TryCreateConvexPath(OfficeClipPath path, OfficeTransform transform,
        CancellationToken cancellationToken, out OfficeTextInkClip clip) {
        clip = default;
        cancellationToken.ThrowIfCancellationRequested();
        if (path.Commands.Count > 512) return false;
        IReadOnlyList<IReadOnlyList<OfficePoint>> contours = OfficeClipPathGeometry.CreateContours(path, points => {
            var mapped = new List<OfficePoint>(points.Count);
            foreach (OfficePoint point in points) {
                cancellationToken.ThrowIfCancellationRequested();
                mapped.Add(transform.TransformPoint(point));
            }
            return mapped;
        });
        if (contours.Count != 1 || contours[0].Count > 513) return false;
        var polygon = new List<OfficePoint>();
        foreach (OfficePoint point in contours[0]) {
            if (!Finite(point.X) || !Finite(point.Y)) return false;
            if (polygon.Count == 0 || point != polygon[polygon.Count - 1]) polygon.Add(point);
        }
        if (polygon.Count > 1 && polygon[0] == polygon[polygon.Count - 1]) polygon.RemoveAt(polygon.Count - 1);
        if (polygon.Count < 3 || polygon.Count > 512) return false;
        // Repeated boundary traversals can cancel under even-odd fill even when
        // every point lies in the same convex hull. They are not simple contours.
        if (new HashSet<OfficePoint>(polygon).Count != polygon.Count) return false;
        double area = 0D;
        for (int i = 1; i < polygon.Count - 1; i++) area += Cross(polygon[0], polygon[i], polygon[i + 1]);
        if (!Finite(area) || Math.Abs(area) < 1e-10D) return false;
        double orientation = Math.Sign(area);
        // Every vertex must lie inside every oriented edge. Checking only adjacent
        // turns would incorrectly accept self-intersecting star polygons.
        for (int edge = 0; edge < polygon.Count; edge++) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficePoint start = polygon[edge], end = polygon[(edge + 1) % polygon.Count];
            foreach (OfficePoint point in polygon) {
                double side = Cross(start, end, point) * orientation;
                if (!Finite(side) || side < -1e-10D) return false;
            }
        }
        clip = new OfficeTextInkClip(polygon, orientation);
        return true;
    }

    private static double Cross(OfficePoint start, OfficePoint end, OfficePoint point) =>
        (end.X - start.X) * (point.Y - start.Y) - (end.Y - start.Y) * (point.X - start.X);
    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}

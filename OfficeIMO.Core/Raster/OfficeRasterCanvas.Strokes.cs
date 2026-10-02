using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    /// <summary>Builds the stroke in local coordinates before applying an affine transform to its outline.</summary>
    internal void StrokeTransformedContours(IReadOnlyList<OfficeFlattenedPathContour> contours, double width,
        OfficeStrokeLineCap cap, OfficeStrokeLineJoin join, double miterLimit, IReadOnlyList<double>? pattern,
        double offset, OfficeTransform transform, Func<double, double, OfficeColor> paint) {
        if (!transform.TryInvert(out OfficeTransform inverse)) {
            // Preserve the scene contract for collapsed transforms: keep the projected
            // centreline visible, while its fill has no two-dimensional area.
            var collapsed = new List<OfficeFlattenedPathContour>(contours.Count);
            foreach (var contour in contours) {
                var ring = new List<OfficePoint>(contour.Points.Count);
                foreach (OfficePoint point in contour.Points) ring.Add(transform.TransformPoint(point));
                collapsed.Add(new OfficeFlattenedPathContour(ring, contour.Closed));
            }
            StrokeContours(collapsed,width,cap,join,miterLimit,pattern,offset,paint);
            return;
        }
        var bounds = inverse.TransformRectangleBounds(0, 0, Width, Height);
        double pixelsPerUnit = Math.Sqrt(transform.M11*transform.M11 + transform.M12*transform.M12 + transform.M21*transform.M21 + transform.M22*transform.M22);
        var outlines = OfficeStrokeGeometry.Create(contours, width, cap, join, miterLimit, pattern, offset, pixelsPerUnit,
            bounds.Left, bounds.Top, bounds.Right, bounds.Bottom, _cancellationToken);
        var transformed = new List<IReadOnlyList<OfficePoint>>(outlines.Count);
        foreach (List<OfficePoint> outline in outlines) {
            for (int i = 0; i < outline.Count; i++) outline[i] = transform.TransformPoint(outline[i]);
            transformed.Add(outline);
        }
        FillContourPaint(transformed, OfficeFillRule.NonZero, paint);
    }

    /// <summary>Paints the union of all stroke pieces, including intersecting subpaths, once.</summary>
    internal void StrokeContours(IReadOnlyList<OfficeFlattenedPathContour> contours, double width,
        OfficeStrokeLineCap cap, OfficeStrokeLineJoin join, double miterLimit,
        IReadOnlyList<double>? dashPattern, double dashOffset, Func<double, double, OfficeColor> paint,
        bool resetDashPatternForEachSegment = false) {
        if (!IsFinite(width) || width <= 0D) return;
        var outlines = OfficeStrokeGeometry.Create(contours, width, cap, join, miterLimit, dashPattern, dashOffset, 1D, 0D, 0D, Width - 1D, Height - 1D, _cancellationToken, resetDashPatternForEachSegment);
        var fillContours = new List<IReadOnlyList<OfficePoint>>(outlines.Count);
        foreach (List<OfficePoint> outline in outlines) fillContours.Add(outline);
        FillContourPaint(fillContours, OfficeFillRule.NonZero, paint);
    }

    private void StrokePolyline(IReadOnlyList<OfficePoint> points, OfficeColor color, double width,
        bool closed = false, IReadOnlyList<double>? pattern = null, double offset = 0D, bool reset = false) {
        if (color.A == 0 || points == null || points.Count == 0) return;
        StrokeContours(new[] { new OfficeFlattenedPathContour(points, closed) }, width,
            OfficeStrokeLineCap.Round, OfficeStrokeLineJoin.Round, 4D, pattern, offset, (_, _) => color, reset);
    }

}

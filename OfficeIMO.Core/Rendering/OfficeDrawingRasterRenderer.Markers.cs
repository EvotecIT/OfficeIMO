using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingRasterRenderer {
    private static void RenderLineMarkers(OfficeRasterCanvas canvas, OfficeShape shape, OfficePoint start, OfficePoint end,
        OfficeColor fallback, OfficeLinearGradient? linear, OfficeRadialGradient? radial, OfficeTransform localToRaster) {
        if (shape.StrokeStartMarker == null && shape.StrokeEndMarker == null) return;
        var bounds = (0D, 0D, shape.Width, shape.Height);
        RenderLineMarker(canvas, shape.StrokeStartMarker, start, end,
            CreateMarkerPaint(fallback, linear, radial, bounds, OfficeTransform.Identity, localToRaster, start), shape, localToRaster);
        RenderLineMarker(canvas, shape.StrokeEndMarker, end, start,
            CreateMarkerPaint(fallback, linear, radial, bounds, OfficeTransform.Identity, localToRaster, end), shape, localToRaster);
    }

    private static void RenderLineMarker(OfficeRasterCanvas canvas, OfficeLineMarker? marker, OfficePoint tip, OfficePoint adjacent,
        Func<double, double, OfficeColor> paint, OfficeShape shape, OfficeTransform localToRaster) {
        // Build both the marker and its outline in source coordinates before applying the shaft's complete transform.
        IReadOnlyList<OfficePoint> contour = OfficeLineMarkerGeometry.CreateContour(marker, tip,
            new OfficePoint(tip.X - adjacent.X, tip.Y - adjacent.Y));
        if (contour.Count < 3) return;
        if (marker!.Kind == OfficeLineMarkerKind.Arrow) {
            canvas.StrokeTransformedContours(new[] { new OfficeFlattenedPathContour(contour, false) },
                shape.StrokeWidth, shape.StrokeLineCap ?? OfficeStrokeLineCap.Round, shape.StrokeLineJoin ?? OfficeStrokeLineJoin.Round,
                shape.StrokeMiterLimit, null, 0D, localToRaster, paint);
        } else {
            var transformed = new List<OfficePoint>(contour.Count);
            foreach (OfficePoint point in contour) transformed.Add(localToRaster.TransformPoint(point));
            canvas.FillContourPaint(new[] { (IReadOnlyList<OfficePoint>)transformed }, OfficeFillRule.NonZero, paint);
        }
    }

    private static void RenderPathMarkers(OfficeRasterCanvas canvas, OfficeShape shape, IReadOnlyList<OfficeFlattenedPathContour> contours,
        OfficeColor fallbackColor, OfficeLinearGradient? linear, OfficeRadialGradient? radial,
        (double X, double Y, double Width, double Height) paintBounds, OfficeTransform contourToLocal, OfficeTransform localToRaster) {
        if (shape.StrokeStartMarker == null && shape.StrokeEndMarker == null) return;
        OfficeFlattenedPathContour? first = null, last = null;
        foreach (OfficeFlattenedPathContour contour in contours) {
            if (contour.Closed || contour.Points.Count < 2) continue;
            first ??= contour;
            last = contour;
        }
        if (first != null) Paint(shape.StrokeStartMarker, first.Points[0], first.Points[1]);
        if (last != null) Paint(shape.StrokeEndMarker, last.Points[last.Points.Count - 1], last.Points[last.Points.Count - 2]);

        void Paint(OfficeLineMarker? marker, OfficePoint tip, OfficePoint adjacent) {
            var paint = CreateMarkerPaint(fallbackColor, linear, radial, paintBounds, contourToLocal, localToRaster, tip);
            RenderLineMarker(canvas, marker, contourToLocal.TransformPoint(tip), contourToLocal.TransformPoint(adjacent), paint, shape, localToRaster);
        }
    }

    private static Func<double, double, OfficeColor> CreateMarkerPaint(OfficeColor fallback, OfficeLinearGradient? linear,
        OfficeRadialGradient? radial, (double X, double Y, double Width, double Height) bounds,
        OfficeTransform contourToLocal, OfficeTransform localToRaster, OfficePoint endpoint) {
        if (linear == null && radial == null) return (_, _) => fallback;
        if (!localToRaster.TryInvert(out var rasterToLocal) || !contourToLocal.TryInvert(out var localToPaint)) {
            // A collapsed transform has no inverse field. Retain the endpoint brush.
            OfficeColor endpointColor = SampleStrokeGradient(linear, radial, bounds.X, bounds.Y, bounds.Width, bounds.Height, endpoint.X, endpoint.Y) ?? fallback;
            return (_, _) => endpointColor;
        }
        OfficeTransform rasterToPaint = rasterToLocal.Then(localToPaint);
        return (x, y) => {
            OfficePoint point = rasterToPaint.TransformPoint(new OfficePoint(x, y));
            return SampleStrokeGradient(linear, radial, bounds.X, bounds.Y, bounds.Width, bounds.Height, point.X, point.Y) ?? fallback;
        };
    }
}

using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingRasterRenderer {
    private static void RenderLineMarkers(OfficeRasterCanvas canvas, OfficeShape shape, OfficePoint start, OfficePoint end,
        OfficeColor startColor, OfficeColor endColor, OfficeTransform localToRaster) {
        RenderLineMarker(canvas, shape.StrokeStartMarker, start, end, startColor, shape, localToRaster);
        RenderLineMarker(canvas, shape.StrokeEndMarker, end, start, endColor, shape, localToRaster);
    }

    private static void RenderLineMarker(OfficeRasterCanvas canvas, OfficeLineMarker? marker, OfficePoint tip, OfficePoint adjacent,
        OfficeColor color, OfficeShape shape, OfficeTransform localToRaster) {
        // Build both the marker and its outline in source coordinates before applying the shaft's complete transform.
        IReadOnlyList<OfficePoint> contour = OfficeLineMarkerGeometry.CreateContour(marker, tip,
            new OfficePoint(tip.X - adjacent.X, tip.Y - adjacent.Y));
        if (contour.Count < 3) return;
        if (marker!.Kind == OfficeLineMarkerKind.Arrow) {
            canvas.StrokeTransformedContours(new[] { new OfficeFlattenedPathContour(contour, false) },
                shape.StrokeWidth, shape.StrokeLineCap ?? OfficeStrokeLineCap.Round, shape.StrokeLineJoin ?? OfficeStrokeLineJoin.Round,
                shape.StrokeMiterLimit, null, 0D, localToRaster, (_, _) => color);
        } else {
            var transformed = new List<OfficePoint>(contour.Count);
            foreach (OfficePoint point in contour) transformed.Add(localToRaster.TransformPoint(point));
            canvas.FillPolygon(transformed, color);
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
            OfficeColor color = SampleStrokeGradient(linear, radial, paintBounds.X, paintBounds.Y, paintBounds.Width, paintBounds.Height, tip.X, tip.Y)
                ?? fallbackColor;
            RenderLineMarker(canvas, marker, contourToLocal.TransformPoint(tip), contourToLocal.TransformPoint(adjacent), color, shape, localToRaster);
        }
    }
}

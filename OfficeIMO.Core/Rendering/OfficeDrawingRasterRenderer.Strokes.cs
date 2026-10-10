using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingRasterRenderer {
    private static void StrokeTransformedPathContours(OfficeRasterCanvas canvas, OfficeDrawingShape drawing,
        IReadOnlyList<OfficeFlattenedPathContour> contours, double scale, OfficeColor? color,
        OfficeLinearGradient? linear, OfficeRadialGradient? radial) {
        OfficeShape shape = drawing.Shape;
        if (shape.StrokeWidth <= 0 || (color == null && linear == null && radial == null)) return;
        OfficeTransform transform = GetRasterTransform(drawing, scale);
        var points = new List<OfficePoint>();
        foreach (OfficeFlattenedPathContour contour in contours) foreach (OfficePoint point in contour.Points) points.Add(transform.TransformPoint(point));
        if (points.Count == 0) return;
        GetPointBounds(points, out double x, out double y, out double width, out double height);
        IReadOnlyList<double>? pattern = shape.StrokeDashArray.Count > 0 ? shape.StrokeDashArray : shape.StrokeDashStyle.GetDashPattern(shape.StrokeWidth);
        bool localPaint = transform.TryInvert(out var inverse);
        canvas.StrokeTransformedContours(contours, shape.StrokeWidth, shape.StrokeLineCap ?? OfficeStrokeLineCap.Round,
            shape.StrokeLineJoin ?? OfficeStrokeLineJoin.Round, shape.StrokeMiterLimit, pattern, shape.StrokeDashOffset, transform,
            (px, py) => {
                if (localPaint) {
                    var point = inverse.TransformPoint(new OfficePoint(px, py));
                    return SampleStrokeGradient(linear, radial, 0D, 0D, shape.Width, shape.Height, point.X, point.Y)
                        ?? color ?? OfficeColor.Transparent;
                }
                return SampleStrokeGradient(linear, radial, x, y, width, height, px, py) ?? color ?? OfficeColor.Transparent;
            });
    }

    private static OfficeTransform GetRasterTransform(OfficeDrawingShape drawing, double scale) =>
        (drawing.Shape.Transform ?? OfficeTransform.Identity)
            .Then(OfficeTransform.Translate(drawing.X, drawing.Y)).Then(OfficeTransform.Scale(scale, scale));
}

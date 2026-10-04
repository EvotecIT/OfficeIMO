using System.Collections.Generic;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    internal static partial class VisioSvgPreviewRasterizer {
        private static OfficeColor GetStrokeFallbackColor(SvgPaint paint) => paint.StrokeGradient?.Stops[0].Color
            ?? paint.StrokeRadialGradient?.Stops[0].Color ?? paint.Stroke;

        private static void StrokeOpenContour(OfficeRasterCanvas canvas, IReadOnlyList<OfficePoint> points, SvgPaint paint, double width, IReadOnlyList<double>? pattern, double dashOffset, SvgTransform transform) =>
            StrokePreviewContours(canvas, new[] { new OfficeFlattenedPathContour(points, false) }, paint, width, pattern, dashOffset, transform);

        private static void StrokeClosedContour(OfficeRasterCanvas canvas, IReadOnlyList<OfficePoint> points, SvgPaint paint, double width, IReadOnlyList<double>? pattern, double dashOffset, SvgTransform transform) =>
            StrokePreviewContours(canvas, new[] { new OfficeFlattenedPathContour(points, true) }, paint, width, pattern, dashOffset, transform);

        private static void StrokePreviewContours(OfficeRasterCanvas canvas, IReadOnlyList<OfficeFlattenedPathContour> contours, SvgPaint paint, double width, IReadOnlyList<double>? pattern, double dashOffset, SvgTransform transform) {
            var points = new List<IReadOnlyList<OfficePoint>>(contours.Count);
            foreach (OfficeFlattenedPathContour contour in contours) points.Add(contour.Points);
            OfficeStrokeLineCap cap = paint.StrokeLineCap == SvgStrokeLineCap.Round ? OfficeStrokeLineCap.Round
                : paint.StrokeLineCap == SvgStrokeLineCap.Square ? OfficeStrokeLineCap.Square : OfficeStrokeLineCap.Butt;
            OfficeStrokeLineJoin join = paint.StrokeLineJoin == SvgStrokeLineJoin.Round ? OfficeStrokeLineJoin.Round
                : paint.StrokeLineJoin == SvgStrokeLineJoin.Bevel ? OfficeStrokeLineJoin.Bevel : OfficeStrokeLineJoin.Miter;
            var sampler = canvas.CreateContourPaint(points, paint.Stroke, paint.StrokeGradient, paint.StrokeRadialGradient);
            OfficeTransform affine = transform.ToOfficeTransform();
            if (affine != OfficeTransform.Identity && !paint.NonScalingStroke && affine.TryInvert(out OfficeTransform inverse)) {
                var local = new List<OfficeFlattenedPathContour>(contours.Count);
                foreach (var contour in contours) {
                    var ring = new List<OfficePoint>(contour.Points.Count);
                    foreach (OfficePoint point in contour.Points) ring.Add(inverse.TransformPoint(point));
                    local.Add(new OfficeFlattenedPathContour(ring, contour.Closed));
                }
                canvas.StrokeTransformedContours(local, paint.StrokeWidth, cap, join, paint.MiterLimit, paint.DashPattern, paint.DashOffset, affine, sampler);
            } else canvas.StrokeContours(contours, width, cap, join, paint.MiterLimit, pattern, dashOffset, sampler);
        }
    }
}

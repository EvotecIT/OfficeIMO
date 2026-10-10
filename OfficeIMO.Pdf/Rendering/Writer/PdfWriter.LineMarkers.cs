using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Projects decorations once for body and running header/footer drawing paths.</summary>
    private static IReadOnlyList<OfficeShape> CreateLineMarkerShapes(OfficeShape source) {
        if (source.StrokeWidth <= 0D || (source.StrokeColor == null && source.StrokeGradient == null && source.StrokeRadialGradient == null)
            || (source.StrokeStartMarker == null && source.StrokeEndMarker == null)) return Array.Empty<OfficeShape>();
        var markers = new List<OfficeShape>(2);
        var contours = OfficeStrokeGeometry.FlattenShape(source, 4D);
        OfficeFlattenedPathContour? first = null, last = null;
        foreach (OfficeFlattenedPathContour contour in contours) {
            if (contour.Closed || contour.Points.Count < 2) continue;
            first ??= contour; last = contour;
        }
        if (first != null) AddMarker(source.StrokeStartMarker, first.Points[0], first.Points[1]);
        if (last != null) AddMarker(source.StrokeEndMarker, last.Points[last.Points.Count - 1], last.Points[last.Points.Count - 2]);
        return markers;

        void AddMarker(OfficeLineMarker? marker, OfficePoint tip, OfficePoint adjacent) {
            IReadOnlyList<OfficePoint> points = OfficeLineMarkerGeometry.CreateContour(marker, tip,
                new OfficePoint(tip.X - adjacent.X, tip.Y - adjacent.Y));
            if (points.Count < 3) return;
            var commands = new List<OfficePathCommand> { OfficePathCommand.MoveTo(points[0]) };
            for (int i = 1; i < points.Count; i++) commands.Add(OfficePathCommand.LineTo(points[i]));
            bool open = marker!.Kind == OfficeLineMarkerKind.Arrow;
            if (!open) commands.Add(OfficePathCommand.Close());
            var paint = OfficeShape.Path(Math.Max(.0001D, source.Width), Math.Max(.0001D, source.Height), commands);
            paint.Transform = source.Transform;
            paint.ClipPath = source.ClipPath?.Clone();
            if (open) {
                paint.StrokeColor = source.StrokeColor;
                paint.StrokeGradient = source.StrokeGradient?.Clone();
                paint.StrokeRadialGradient = source.StrokeRadialGradient?.Clone();
                paint.StrokeWidth = source.StrokeWidth;
                paint.StrokeOpacity = source.StrokeOpacity;
                paint.StrokeLineCap = source.StrokeLineCap; paint.StrokeLineJoin = source.StrokeLineJoin;
                paint.StrokeMiterLimit = source.StrokeMiterLimit;
            } else {
                paint.FillColor = source.StrokeColor;
                paint.FillGradient = source.StrokeGradient?.Clone();
                paint.FillRadialGradient = source.StrokeRadialGradient?.Clone();
                paint.FillOpacity = source.StrokeOpacity;
            }
            markers.Add(paint);
        }
    }

    private sealed partial class LayoutContext {
        private void DrawShapeMarkersAt(OfficeShape source, double x, double bottomY) {
            foreach (OfficeShape marker in CreateLineMarkerShapes(source))
                DrawShapeGeometryAt(marker, x, bottomY + source.Height - marker.Height);
        }
    }

    private static void DrawHeaderFooterShapeMarkersAt(System.Text.StringBuilder sb, LayoutResult.Page page, OfficeShape source, double x, double bottomY) {
        foreach (OfficeShape marker in CreateLineMarkerShapes(source))
            DrawHeaderFooterShapeGeometryAt(sb, page, marker, x, bottomY + source.Height - marker.Height);
    }
}

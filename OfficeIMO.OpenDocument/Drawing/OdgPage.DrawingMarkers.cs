using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgPage {
    // Isolated affine Drawing groups rasterize their declared canvas. Include paint outside the source shape box.
    private static OfficeDrawing ExpandStrokeCanvas(OfficeDrawing local, OfficeShape? shape, out double leftPadding, out double topPadding, bool includeCanvas = true) {
        double margin = shape != null && shape.StrokeColor.HasValue ? shape.StrokeWidth * Math.Max(1, shape.StrokeMiterLimit) / 2 : 0;
        double left = includeCanvas ? -margin : double.PositiveInfinity, top = left;
        double right = includeCanvas ? local.Width + margin : double.NegativeInfinity, bottom = includeCanvas ? local.Height + margin : right;
        if (shape?.Kind == OfficeShapeKind.Path) {
            var path = OfficePathGeometry.Bounds(shape.PathCommands);
            left = Math.Min(left, path.Left - margin); top = Math.Min(top, path.Top - margin);
            right = Math.Max(right, path.Right + margin); bottom = Math.Max(bottom, path.Bottom + margin);
        }
        foreach (OfficeDrawingEffectGroup marker in local.Elements.OfType<OfficeDrawingEffectGroup>()) {
            var box = marker.Transform.TransformRectangleBounds(0, 0, marker.InnerDrawing.Width, marker.InnerDrawing.Height);
            left = Math.Min(left, box.Left); top = Math.Min(top, box.Top); right = Math.Max(right, box.Right); bottom = Math.Max(bottom, box.Bottom);
        }
        leftPadding = -left; topPadding = -top;
        var expanded = new OfficeDrawing(right - left, bottom - top);
        CopyDrawingResources(local, expanded);
        if (includeCanvas) expanded.AddDrawing(local, -left, -top);
        else expanded.AddDrawingForClippedRendering(local, -left, -top, null);
        return expanded;
    }
    private static bool SuppressedMarkerWidth(OdgShape source, bool start) {
        try {
            OdfLength? width = start ? source.StrokeStartMarkerWidth : source.StrokeEndMarkerWidth;
            return width.HasValue && OdfShape.ValidateMarkerWidth(width.Value) == 0;
        }
        catch (ArgumentException) { return false; } // Active malformed metrics remain conversion losses.
        catch (FormatException) { return false; }
    }

    private static MarkerPaint? ResolveMarker(OdgShape source, OfficeShape stroke, OdfConversionReport report,
        string position, string? name) {
        if (name == null) return null;
        try {
            OdfLength? widthValue = position == "start" ? source.StrokeStartMarkerWidth : source.StrokeEndMarkerWidth;
            if (!widthValue.HasValue) throw new NotSupportedException("Marker width is unspecified; no consumer-dependent default is assumed.");
            double width = OdfShape.ValidateMarkerWidth(widthValue.Value);
            if (width == 0) return null;
            OdfMarker marker = source.Document.Styles.FindMarker(name) ?? throw new InvalidDataException("Missing marker '" + name + "'.");
            OdfMarkerGeometry geometry = marker.Geometry;
            bool center = (position == "start" ? source.StrokeStartMarkerCentered : source.StrokeEndMarkerCentered) ?? false;
            var bounds = geometry.Bounds;
            double scale = width / (bounds.Right - bounds.Left), height = (bounds.Bottom - bounds.Top)*scale;
            OfficeShape painted = OfficeShape.Path(width, height, geometry.Commands.Select(command => command.Translate(bounds.Left, bounds.Top).Scale(scale, scale)));
            painted.FillColor = stroke.StrokeColor; painted.StrokeColor = null;
            painted.FillRule = OfficeFillRule.NonZero;
            return new MarkerPaint(painted, width, height, center);
        } catch (Exception exception) when (exception is FormatException || exception is ArgumentException || exception is NotSupportedException || exception is InvalidDataException || exception is OverflowException) {
            report.Add("shape:" + source.Name + ":marker-" + position, OdfConversionMappingStatus.Unsupported, message: exception.Message);
            return null;
        }
    }

    private static void PlaceMarker(OfficeDrawing paint, MarkerPaint? marker, OfficePoint endpoint, OfficePoint tangent, OfficePoint consumedPoint) {
        if (marker == null) return;
        var inward = new OfficePoint(consumedPoint.X - endpoint.X, consumedPoint.Y - endpoint.Y);
        if (inward.X == 0 && inward.Y == 0) inward = tangent;
        double angle = Math.Atan2(inward.Y, inward.X)*180D/Math.PI - 90;
        var scene = new OfficeDrawing(marker.Width, marker.Height); scene.AddShape(marker.Shape, 0, 0);
        paint.AddEffectDrawing(scene, OfficeTransform.Translate(-marker.Width/2, -marker.Dock)
            .Then(OfficeTransform.RotateDegrees(angle)).Then(OfficeTransform.Translate(endpoint.X, endpoint.Y)));
    }

    private sealed class MarkerPaint {
        internal OfficeShape Shape { get; }
        internal double Width { get; }
        internal double Height { get; }
        internal double Dock { get; }
        internal double Consumed { get; }
        internal double Trim => Math.Max(0, Consumed - Width/15);
        internal MarkerPaint(OfficeShape shape, double width, double height, bool center) {
            Shape = shape; Width = width; Height = height; Dock = center ? height/2 : 0;
            Consumed = CenterlineExit(shape.PathCommands, width/2, Dock, height - Dock);
        }
    }

    // Marker polygon vertices determine a notched back's centerline exit, matching native docking.
    // Curved marker outlines keep their exact geometry; this query uses their endpoint polygon.
    private static double CenterlineExit(IReadOnlyList<OfficePathCommand> commands, double x, double dock, double fallback) {
        var contours = new List<List<OfficePoint>>(); List<OfficePoint>? current = null;
        foreach (OfficePathCommand command in commands) {
            if (command.Kind == OfficePathCommandKind.MoveTo) { current = new List<OfficePoint>(); contours.Add(current); }
            if (command.Kind != OfficePathCommandKind.Close && current != null && (current.Count == 0 || current[current.Count - 1] != command.Point)) current.Add(command.Point);
        }
        var crossings = new List<double>();
        foreach (List<OfficePoint> polygon in contours) {
            if (polygon.Count > 1 && polygon[0] == polygon[polygon.Count - 1]) polygon.RemoveAt(polygon.Count - 1);
            if (polygon.Count < 2) continue;
            for (int i = 0; i < polygon.Count; i++) {
                OfficePoint a = polygon[i], b = polygon[(i + 1)%polygon.Count];
                double ax = a.X - x, bx = b.X - x;
                if ((ax < 0 && bx > 0) || (ax > 0 && bx < 0)) crossings.Add(a.Y + (b.Y - a.Y)*(-ax/(bx - ax)) - dock);
                else if (ax == 0 && bx != 0) {
                    double previous = polygon[(i + polygon.Count - 1)%polygon.Count].X - x;
                    if ((previous < 0 && bx > 0) || (previous > 0 && bx < 0)) crossings.Add(a.Y - dock);
                }
            }
        }
        crossings.Sort();
        if (crossings.Count(y => y < 1e-9)%2 == 1) {
            foreach (double y in crossings) if (y > 1e-9) return Math.Min(fallback, y);
        }
        return fallback;
    }
}

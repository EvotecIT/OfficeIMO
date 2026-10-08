using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgPage {
    // Draw has one closure state for an entire path, unlike SVG's independent subpath closure.
    private static bool ProjectStrokeContours(OdgShape source, OfficeShape stroke, OfficeDrawing local, OdfConversionReport report) {
        string feature = "shape:" + source.Name;
        IReadOnlyList<OfficePathContour>? contours = stroke.Kind == OfficeShapeKind.Path ? OfficePathContour.Split(stroke.PathCommands) : null;
        bool closed = stroke.Kind != OfficeShapeKind.Line && (contours == null || contours.All(contour => contour.IsClosed));
        if (!closed && HasFill(stroke)) {
            ClearFill(stroke);
            report.Add(feature + ":inactive-fill", OdfConversionMappingStatus.Converted,
                message: "Draw open paths do not paint their declared fill. The original fill and path XML remain preserved.");
        }
        if (!stroke.StrokeColor.HasValue) return false;
        string? startName = source.StrokeStartMarkerName, endName = source.StrokeEndMarkerName;
        if (closed) {
            if (startName != null || endName != null)
                report.Add(feature + ":markers", OdfConversionMappingStatus.Converted, message: "Closed Draw shapes do not paint stroke markers; their bindings remain preserved.");
            startName = endName = null;
        } else {
            if (startName != null && SuppressedMarkerWidth(source, true)) startName = null;
            if (endName != null && SuppressedMarkerWidth(source, false)) endName = null;
        }
        if ((contours == null || contours.Count == 1) && startName == null && endName == null) return false;
        var paths = contours != null ? contours.Select(contour => closed ? contour.Commands : contour.OpenWithClosingEdge()).ToArray() :
            new IReadOnlyList<OfficePathCommand>[] { new[] { OfficePathCommand.MoveTo(stroke.Points[0]), OfficePathCommand.LineTo(stroke.Points[1]) } };
        bool drawable = paths.Any(path => OfficePathGeometry.TryOpenEndpoints(path, out _, out _, out _, out _));
        MarkerPaint? first = drawable ? ResolveMarker(source, stroke, report, "start", startName) : null;
        MarkerPaint? last = drawable ? ResolveMarker(source, stroke, report, "end", endName) : null;
        var measures = new OfficePathMeasure?[paths.Length];
        if (first != null || last != null) {
            int samples = 0;
            try {
                // Resolve the whole shape before painting; the budget cannot reset at every contour.
                for (int i = 0; i < paths.Length; i++)
                    if (OfficePathGeometry.TryOpenEndpoints(paths[i], out _, out _, out _, out _))
                        measures[i] = new OfficePathMeasure(paths[i], ref samples);
            } catch (NotSupportedException exception) {
                report.Add(feature + ":markers", OdfConversionMappingStatus.Unsupported, message: exception.Message);
                first = last = null;
                Array.Clear(measures, 0, measures.Length);
            }
        }
        if (HasFill(stroke)) {
            OfficeShape fill = stroke.Clone(); fill.StrokeColor = null; fill.StrokeGradient = null; fill.StrokeRadialGradient = null;
            local.AddShape(fill, 0, 0);
        }
        int paintedContours = 0;
        for (int i = 0; i < paths.Length; i++) {
            IReadOnlyList<OfficePathCommand> path = paths[i];
            if (!path.Any(command => command.Kind is OfficePathCommandKind.LineTo or OfficePathCommandKind.QuadraticBezierTo or OfficePathCommandKind.CubicBezierTo)) continue;
            var paint = new OfficeDrawing(local.Width, local.Height);
            OfficePathMeasure? measure = measures[i];
            OfficeShape? shaft = null;
            IReadOnlyList<OfficePathCommand> shortened = measure?.Slice(first?.Trim ?? 0, measure.Length - (last?.Trim ?? 0)) ?? path;
            if (shortened.Count > 0) {
                shaft = stroke.CloneWithPath(local.Width, local.Height, shortened);
                ClearFill(shaft);
                if (measure == null) {
                    // A stroke without markers already applies opacity once to its own contour.
                    local.AddShape(shaft, 0, 0); continue;
                }
                shaft.StrokeOpacity = 1;
                paint.AddShape(shaft, 0, 0);
            }
            if (measure != null && OfficePathGeometry.TryOpenEndpoints(path, out var start, out var forward, out var end, out var backward)) {
                PlaceMarker(paint, first, start, forward, measure.PointAtLength(first?.Consumed ?? 0));
                PlaceMarker(paint, last, end, backward, measure.PointAtLength(measure.Length - (last?.Consumed ?? 0)));
                paintedContours++;
            }
            double opacity = stroke.StrokeOpacity ?? 1;
            if (opacity < 1) {
                // Shaft and markers share opacity, but distinct contours accumulate transparency.
                paint = ExpandStrokeCanvas(paint, shaft, out double left, out double top, includeCanvas: false);
                local.AddEffectDrawing(paint, OfficeTransform.Translate(-left, -top), opacity);
            } else local.AddDrawing(paint, 0, 0);
        }
        foreach (string position in new[] { "start", "end" }) {
            if ((position == "start" ? first : last) != null && paintedContours > 0)
                report.Add(feature + ":marker-" + position, OdfConversionMappingStatus.Approximated,
                    message: "Markers are projected on " + paintedContours + " drawable contours with width/15 overlap and consumed-path chord orientation. Stroke opacity covers each contour's shaft and markers together.");
        }
        if (!drawable && (startName != null || endName != null))
            report.Add(feature + ":markers", OdfConversionMappingStatus.Converted, message: "Degenerate contours do not paint stroke markers; their bindings remain preserved.");
        return true;
    }

    private static bool HasFill(OfficeShape shape) => shape.FillColor.HasValue || shape.FillGradient != null || shape.FillRadialGradient != null;
    private static void ClearFill(OfficeShape shape) { shape.FillColor = null; shape.FillGradient = null; shape.FillRadialGradient = null; }
}

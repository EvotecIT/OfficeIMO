using System;
using System.Text;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingSvgExporter {
    private static bool UsesSharedMarkerField(OfficeShape shape) {
        if (shape.Kind is not (OfficeShapeKind.Line or OfficeShapeKind.Path)
            || (shape.StrokeStartMarker == null && shape.StrokeEndMarker == null)) return false;
        if (!OfficeLineMarkerGeometry.TryEndpoints(shape, out var start, out var startDirection, out var end, out var endDirection)) return false;
        return OfficeLineMarkerGeometry.CreateContour(shape.StrokeStartMarker, start, startDirection).Count >= 3
            || OfficeLineMarkerGeometry.CreateContour(shape.StrokeEndMarker, end, endDirection).Count >= 3;
    }

    private static void AppendPathMarkers(StringBuilder sb, OfficeShape shape, double originX, double originY, string transform, string? strokeGradientId) {
        if (shape.StrokeStartMarker == null && shape.StrokeEndMarker == null) return;
        if (!OfficeLineMarkerGeometry.TryEndpoints(shape, out var start, out var startDirection, out var end, out var endDirection)) return;
        AppendLineMarker(sb, shape.StrokeStartMarker, new(start.X + originX, start.Y + originY), startDirection, shape, transform, strokeGradientId);
        AppendLineMarker(sb, shape.StrokeEndMarker, new(end.X + originX, end.Y + originY), endDirection, shape, transform, strokeGradientId);
    }

    private static OfficeTransform GetStrokeFieldCoordinates(OfficeDrawingShape drawing) {
        OfficeShape shape = drawing.Shape;
        bool local = shape.ClipPath != null || HasNonIdentityTransform(shape.Transform);
        return new OfficeTransform(Math.Max(.0001D, shape.Width), 0D, 0D, Math.Max(.0001D, shape.Height),
            local ? 0D : drawing.X, local ? 0D : drawing.Y);
    }

    private static void AppendStrokeLinearPaintDefinition(StringBuilder output, string id, OfficeLinearGradient gradient,
        OfficeDrawingShape drawing) {
        if (UsesSharedMarkerField(drawing.Shape)) output.AppendLinearGradientFieldDefinition(id, gradient, GetStrokeFieldCoordinates(drawing));
        else output.AppendLinearGradientDefinition(id, gradient);
    }

    private static void AppendStrokeRadialPaintDefinition(StringBuilder output, string id, OfficeRadialGradient gradient,
        OfficeDrawingShape drawing, CancellationToken token) {
        if (UsesSharedMarkerField(drawing.Shape) && gradient.OutsideColor == null)
            output.AppendRadialGradientFieldDefinition(id, gradient.TransformCoordinates(GetStrokeFieldCoordinates(drawing)));
        else AppendRadialPaintDefinition(output, id, gradient, drawing, token);
    }
}

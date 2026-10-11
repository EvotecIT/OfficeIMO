using System;
using System.Text;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingSvgExporter {
    private static bool UsesSharedMarkerField(OfficeShape shape) {
        if (shape.Kind is not (OfficeShapeKind.Line or OfficeShapeKind.Path)
            || (shape.StrokeStartMarker == null && shape.StrokeEndMarker == null)) return false;
        OfficeFlattenedPathContour? first = null, last = null;
        foreach (var contour in OfficeStrokeGeometry.FlattenShape(shape, 1D)) {
            if (contour.Closed || contour.Points.Count < 2) continue;
            first ??= contour; last = contour;
        }
        bool Paints(OfficeLineMarker? marker, OfficePoint tip, OfficePoint adjacent) =>
            OfficeLineMarkerGeometry.CreateContour(marker, tip, new OfficePoint(tip.X - adjacent.X, tip.Y - adjacent.Y)).Count >= 3;
        return first != null && Paints(shape.StrokeStartMarker, first.Points[0], first.Points[1])
            || last != null && Paints(shape.StrokeEndMarker, last.Points[last.Points.Count - 1], last.Points[last.Points.Count - 2]);
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

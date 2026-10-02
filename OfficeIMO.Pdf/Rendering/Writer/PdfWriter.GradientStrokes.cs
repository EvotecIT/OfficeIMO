using System.Collections.Generic;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Projects a stroke to fill geometry for the shared native gradient writer.</summary>
    private static OfficeShape? CreateGradientStrokeShape(OfficeShape shape) {
        var contours = OfficeStrokeGeometry.FlattenShape(shape, 4D);
        IReadOnlyList<double>? pattern = shape.StrokeDashArray.Count > 0
            ? shape.StrokeDashArray : shape.StrokeDashStyle.GetDashPattern(shape.StrokeWidth);
        var outlines = OfficeStrokeGeometry.Create(contours, shape.StrokeWidth,
            shape.StrokeLineCap ?? OfficeStrokeLineCap.Round, shape.StrokeLineJoin ?? OfficeStrokeLineJoin.Round,
            shape.StrokeMiterLimit, pattern, shape.StrokeDashOffset, 4D,
            double.NegativeInfinity, double.NegativeInfinity, double.PositiveInfinity, double.PositiveInfinity);
        if (outlines.Count == 0) return null;
        var commands = new List<OfficePathCommand>();
        foreach (var ring in outlines) {
            if (ring.Count < 3) continue;
            commands.Add(OfficePathCommand.MoveTo(ring[0]));
            for (int i = 1; i < ring.Count; i++) commands.Add(OfficePathCommand.LineTo(ring[i]));
            commands.Add(OfficePathCommand.Close());
        }
        // Line bounds can have a zero axis. Keep their coordinate origin while satisfying
        // the path canvas contract; the outline itself carries the visible stroke extent.
        OfficeShape stroke = OfficeShape.Path(Math.Max(.0001D, shape.Width), Math.Max(.0001D, shape.Height), commands);
        stroke.FillColor = null;
        stroke.FillGradient = shape.StrokeGradient;
        stroke.FillRadialGradient = shape.StrokeRadialGradient;
        stroke.FillOpacity = shape.StrokeOpacity;
        stroke.StrokeWidth = 0D;
        stroke.StrokeColor = null;
        stroke.FillRule = OfficeFillRule.NonZero;
        stroke.ClipPath = shape.ClipPath;
        OfficeTransform transform = shape.Transform ?? OfficeTransform.Identity;
        stroke.Transform = new OfficeTransform(transform.M11, transform.M12, transform.M21, transform.M22,
            transform.OffsetX, transform.OffsetY + stroke.Height - shape.Height);
        return stroke;
    }
}

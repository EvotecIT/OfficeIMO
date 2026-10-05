using System;
using System.Linq;
using System.Text;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingSvgExporter {
    private static void AppendRadialPaintDefinition(StringBuilder builder, string id, OfficeRadialGradient gradient,
        OfficeDrawingShape drawingShape, CancellationToken cancellationToken) {
        if (gradient.OutsideColor == null) { builder.AppendRadialGradientDefinition(id, gradient); return; }
        OfficeShape shape = drawingShape.Shape;
        if (!(shape.Width > 0D) || !(shape.Height > 0D))
            throw new NotSupportedException("Native radial SVG paint requires a nonempty shape canvas.");
        double left = 0D, top = 0D, right = shape.Width, bottom = shape.Height;
        foreach (OfficePathCommand command in shape.PathCommands) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeSvgDrawingReader.IncludeCommandBounds(command, ref left, ref top, ref right, ref bottom);
        }
        foreach (OfficePoint point in shape.Points) {
            cancellationToken.ThrowIfCancellationRequested();
            left = Math.Min(left, point.X); top = Math.Min(top, point.Y);
            right = Math.Max(right, point.X); bottom = Math.Max(bottom, point.Y);
        }
        // Include curve control hulls, joins, caps and independently sized markers.
        // The single tile covers every painted point; it must never visibly repeat.
        double margin = Math.Max(1D, shape.StrokeWidth * Math.Max(1D, shape.StrokeMiterLimit));
        foreach (OfficeLineMarker? marker in new[] { shape.StrokeStartMarker, shape.StrokeEndMarker })
            if (marker != null) margin = Math.Max(margin, marker.Width + marker.Length);
        bool local = shape.ClipPath != null || HasNonIdentityTransform(shape.Transform);
        double x = local ? 0D : drawingShape.X, y = local ? 0D : drawingShape.Y;
        left += x - margin; top += y - margin; right += x + margin; bottom += y + margin;
        OfficeRadialGradient field = gradient.TransformCoordinates(new OfficeTransform(shape.Width, 0D, 0D, shape.Height, x, y));
        AppendNativeRadialPattern(builder, id, field, left, top, right - left, bottom - top);
    }

    private static void AppendNativeRadialPattern(StringBuilder builder, string id, OfficeRadialGradient field,
        double left, double top, double width, double height) {
        var colors = field.WithStops(field.Stops.Select(stop => new OfficeGradientStop(stop.Offset,
            OfficeColor.FromRgb(stop.Color.R, stop.Color.G, stop.Color.B))).ToArray());
        var alpha = field.WithStops(field.Stops.Select(stop => new OfficeGradientStop(stop.Offset,
            OfficeColor.FromRgb(stop.Color.A, stop.Color.A, stop.Color.A))).ToArray());
        builder.AppendNativeRadialPatternDefinition(id, left, top, width, height,
            OfficeSvgFormatting.ToCssColor(colors.Stops[0].Color), OfficeSvgFormatting.ToCssColor(alpha.Stops[0].Color),
            (output, fieldId) => output.AppendRadialGradientFieldDefinition(fieldId, colors),
            (output, fieldId) => output.AppendRadialGradientFieldDefinition(fieldId, alpha));
    }
}

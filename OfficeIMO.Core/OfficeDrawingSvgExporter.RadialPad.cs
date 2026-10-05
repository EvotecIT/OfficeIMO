using System;
using System.Linq;
using System.Globalization;
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
        string N(double value) {
            if (double.IsNaN(value) || double.IsInfinity(value)) throw new NotSupportedException("Native radial SVG paint bounds must be finite.");
            return value.ToString("R", CultureInfo.InvariantCulture);
        }
        string region = " x=\"" + N(left) + "\" y=\"" + N(top) + "\" width=\"" + N(width) + "\" height=\"" + N(height) + "\"";
        string colorId = id + "-color", alphaId = id + "-alpha", maskId = id + "-mask";
        // SVG2 shrinking circles select the native first intersection. SVG leaves
        // the outside cone transparent, so compose opaque RGB and alpha separately,
        // just as the PDF owner does. Painting translucent layers over one another
        // would incorrectly compound the outside endpoint alpha inside the cone.
        var colors = field.WithStops(field.Stops.Select(stop => new OfficeGradientStop(stop.Offset,
            OfficeColor.FromRgb(stop.Color.R, stop.Color.G, stop.Color.B))).ToArray());
        var alpha = field.WithStops(field.Stops.Select(stop => new OfficeGradientStop(stop.Offset,
            OfficeColor.FromRgb(stop.Color.A, stop.Color.A, stop.Color.A))).ToArray());
        builder.Append("<defs><pattern").AppendAttribute("id", id).Append(region)
            .Append(" patternUnits=\"userSpaceOnUse\" patternContentUnits=\"userSpaceOnUse\"><g transform=\"translate(")
            .Append(N(-left)).Append(' ').Append(N(-top)).Append(")\">");
        builder.AppendRadialGradientFieldDefinition(colorId, colors);
        builder.AppendRadialGradientFieldDefinition(alphaId, alpha);
        builder.Append("<defs><mask").AppendAttribute("id", maskId).Append(region)
            .Append(" maskUnits=\"userSpaceOnUse\" maskContentUnits=\"userSpaceOnUse\" style=\"mask-type:luminance\">");
        Rect(OfficeSvgFormatting.ToCssColor(alpha.Stops[0].Color)); Rect("url(#" + alphaId + ")");
        builder.Append("</mask></defs><g").AppendAttribute("mask", "url(#" + maskId + ")").Append('>');
        Rect(OfficeSvgFormatting.ToCssColor(colors.Stops[0].Color)); Rect("url(#" + colorId + ")");
        builder.Append("</g></g></pattern></defs>");
        void Rect(string fill) => builder.Append("<rect").Append(region).AppendAttribute("fill", fill).Append("/>");
    }
}

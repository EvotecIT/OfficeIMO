using System;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgFormatting {
    /// <summary>
    /// Appends a reusable SVG linear-gradient definition.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="id">Gradient identifier.</param>
    /// <param name="gradient">Gradient definition.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendLinearGradientDefinition(this StringBuilder builder, string id, OfficeLinearGradient gradient) =>
        AppendLinearGradientDefinitionCore(builder, id, gradient, null);

    internal static StringBuilder AppendLinearGradientFieldDefinition(this StringBuilder builder, string id, OfficeLinearGradient gradient,
        OfficeTransform coordinates) => AppendLinearGradientDefinitionCore(builder, id, gradient, coordinates);

    private static StringBuilder AppendLinearGradientDefinitionCore(StringBuilder builder, string id, OfficeLinearGradient gradient,
        OfficeTransform? coordinates) {
        if (gradient == null) {
            throw new ArgumentNullException(nameof(gradient));
        }

        double multiplier = coordinates.HasValue ? 1D : 100D;
        string unit = coordinates.HasValue ? "" : "%";
        string FormatCoordinate(double value) => coordinates.HasValue ? FormatPreciseNumber(value) : FormatNumber(value);
        builder.Append("<defs><linearGradient id=\"")
            .Append(Escape(id))
            .Append("\" x1=\"")
            .Append(FormatCoordinate(gradient.StartX * multiplier))
            .Append(unit).Append("\" y1=\"")
            .Append(FormatCoordinate(gradient.StartY * multiplier))
            .Append(unit).Append("\" x2=\"")
            .Append(FormatCoordinate(gradient.EndX * multiplier))
            .Append(unit).Append("\" y2=\"")
            .Append(FormatCoordinate(gradient.EndY * multiplier))
            .Append(unit).Append('"');
        if (coordinates.HasValue) {
            OfficeTransform field = coordinates.Value;
            builder.Append(" gradientUnits=\"userSpaceOnUse\" gradientTransform=\"matrix(")
                .Append(FormatPreciseNumber(field.M11)).Append(' ').Append(FormatPreciseNumber(field.M12)).Append(' ')
                .Append(FormatPreciseNumber(field.M21)).Append(' ').Append(FormatPreciseNumber(field.M22)).Append(' ')
                .Append(FormatPreciseNumber(field.OffsetX)).Append(' ').Append(FormatPreciseNumber(field.OffsetY)).Append(")\"");
        }
        if (gradient.ColorInterpolation == OfficeGradientColorInterpolation.LinearRgb) builder.Append(" color-interpolation=\"linearRGB\"");
        builder.Append('>');

        for (int i = 0; i < gradient.Stops.Count; i++) {
            OfficeGradientStop stop = gradient.Stops[i];
            builder.Append("<stop offset=\"")
                .Append(FormatNumber(stop.Offset * 100D))
                .Append("%\" stop-color=\"")
                .Append(ToCssColor(stop.Color))
                .Append('"');

            double opacity = ToOpacity(stop.Color);
            if (opacity < 1D) {
                builder.AppendNumberAttribute("stop-opacity", opacity);
            }

            builder.Append("/>");
        }

        builder.Append("</linearGradient></defs>");
        return builder;
    }

    /// <summary>
    /// Appends a reusable SVG radial-gradient definition.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="id">Gradient identifier.</param>
    /// <param name="gradient">Gradient definition.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendRadialGradientDefinition(this StringBuilder builder, string id, OfficeRadialGradient gradient) =>
        AppendRadialGradientDefinitionCore(builder, id, gradient, false);

    // Native paint composition supplies the missing outside-cone paint and alpha.
    // Its fields use explicit user coordinates, independent of each paint rectangle.
    internal static StringBuilder AppendRadialGradientFieldDefinition(this StringBuilder builder, string id, OfficeRadialGradient gradient) =>
        AppendRadialGradientDefinitionCore(builder, id, gradient, true);

    private static StringBuilder AppendRadialGradientDefinitionCore(StringBuilder builder, string id, OfficeRadialGradient gradient, bool userSpace) {
        if (gradient == null) {
            throw new ArgumentNullException(nameof(gradient));
        }

        if (gradient.OutsideColor != null && !userSpace) {
            throw new NotSupportedException("Native radial boundary/exterior Pad fields cannot be represented by ordinary SVG without loss.");
        }

        bool elliptical = !gradient.EndRadiusX.Equals(gradient.EndRadiusY);
        double endX = elliptical ? 0D : gradient.EndX;
        double endY = elliptical ? 0D : gradient.EndY;
        double endRadius = elliptical ? 1D : gradient.EndRadius;
        double startX = elliptical ? (gradient.StartX - gradient.EndX) / gradient.EndRadiusX : gradient.StartX;
        double startY = elliptical ? (gradient.StartY - gradient.EndY) / gradient.EndRadiusY : gradient.StartY;
        double startRadius = elliptical ? gradient.StartRadiusX / gradient.EndRadiusX : gradient.StartRadius;

        double multiplier = userSpace ? 1D : 100D;
        string unit = userSpace ? "" : "%";
        builder.Append("<defs><radialGradient id=\"")
            .Append(Escape(id))
            .Append("\" cx=\"")
            .Append(FormatPreciseNumber(endX * multiplier))
            .Append(unit).Append("\" cy=\"")
            .Append(FormatPreciseNumber(endY * multiplier))
            .Append(unit).Append("\" r=\"")
            .Append(FormatPreciseNumber(endRadius * multiplier))
            .Append(unit).Append("\" fx=\"")
            .Append(FormatPreciseNumber(startX * multiplier))
            .Append(unit).Append("\" fy=\"")
            .Append(FormatPreciseNumber(startY * multiplier))
            .Append(unit)
            .Append('"');

        if (gradient.ColorInterpolation == OfficeGradientColorInterpolation.LinearRgb) builder.Append(" color-interpolation=\"linearRGB\"");
        if (gradient.SpreadMode != OfficeGradientSpreadMode.Pad) builder.Append(" spreadMethod=\"").Append(gradient.SpreadMode == OfficeGradientSpreadMode.Repeat ? "repeat" : "reflect").Append('"');
        if (userSpace) builder.Append(" gradientUnits=\"userSpaceOnUse\"");

        var coordinates = (elliptical ? new OfficeTransform(gradient.EndRadiusX, 0D, 0D, gradient.EndRadiusY, gradient.EndX, gradient.EndY)
            : OfficeTransform.Identity).Then(gradient.CoordinateTransform);
        if (coordinates != OfficeTransform.Identity) {
            builder.Append(" gradientTransform=\"matrix(")
                .Append(FormatPreciseNumber(coordinates.M11)).Append(' ').Append(FormatPreciseNumber(coordinates.M12)).Append(' ')
                .Append(FormatPreciseNumber(coordinates.M21)).Append(' ').Append(FormatPreciseNumber(coordinates.M22)).Append(' ')
                .Append(FormatPreciseNumber(coordinates.OffsetX)).Append(' ').Append(FormatPreciseNumber(coordinates.OffsetY)).Append(")\"");
        }

        if (startRadius > 0D) {
            builder.Append(" fr=\"")
                .Append(FormatPreciseNumber(startRadius * multiplier))
                .Append(unit).Append("\"");
        }

        builder.Append('>');

        for (int i = 0; i < gradient.Stops.Count; i++) {
            OfficeGradientStop stop = gradient.Stops[i];
            builder.Append("<stop offset=\"")
                .Append(FormatPreciseNumber(stop.Offset * 100D))
                .Append("%\" stop-color=\"")
                .Append(ToCssColor(stop.Color))
                .Append('"');

            double opacity = ToOpacity(stop.Color);
            if (opacity < 1D) {
                builder.AppendNumberAttribute("stop-opacity", opacity);
            }

            builder.Append("/>");
        }

        builder.Append("</radialGradient></defs>");
        return builder;
    }

}

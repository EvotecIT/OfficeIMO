using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static class PdfVisualResourceDictionaryBuilder {
    internal static string BuildExtGStateObject(double fillOpacity, double strokeOpacity) {
        ValidateOpacity(fillOpacity, nameof(fillOpacity));
        ValidateOpacity(strokeOpacity, nameof(strokeOpacity));

        return "<< /Type /ExtGState /ca " +
            FormatNumber(fillOpacity) +
            " /CA " +
            FormatNumber(strokeOpacity) +
            " >>\n";
    }

    internal static string BuildAxialShadingObject(
        double x0,
        double y0,
        double x1,
        double y1,
        OfficeColor startColor,
        OfficeColor endColor) => BuildAxialShadingObject(
            x0,
            y0,
            x1,
            y1,
            new[] { new OfficeGradientStop(0D, startColor), new OfficeGradientStop(1D, endColor) });

    internal static string BuildAxialShadingObject(
        double x0,
        double y0,
        double x1,
        double y1,
        IReadOnlyList<OfficeGradientStop> stops) {
        ValidateFinite(x0, nameof(x0));
        ValidateFinite(y0, nameof(y0));
        ValidateFinite(x1, nameof(x1));
        ValidateFinite(y1, nameof(y1));
        ValidateStops(stops);

        return
            "<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [" +
            FormatNumber(x0) + " " + FormatNumber(y0) + " " + FormatNumber(x1) + " " + FormatNumber(y1) +
            "] /Function " + BuildGradientFunction(stops) + " /Extend [true true] >>\n";
    }

    private static string BuildGradientFunction(IReadOnlyList<OfficeGradientStop> stops) {
        if (stops.Count == 2) return BuildInterpolationFunction(stops[0].Color, stops[1].Color);

        var builder = new System.Text.StringBuilder("<< /FunctionType 3 /Domain [0 1] /Functions [");
        for (int index = 1; index < stops.Count; index++) {
            if (index > 1) builder.Append(' ');
            builder.Append(BuildInterpolationFunction(stops[index - 1].Color, stops[index].Color));
        }

        builder.Append("] /Bounds [");
        for (int index = 1; index < stops.Count - 1; index++) {
            if (index > 1) builder.Append(' ');
            builder.Append(FormatNumber(stops[index].Offset));
        }

        builder.Append("] /Encode [");
        for (int index = 1; index < stops.Count; index++) {
            if (index > 1) builder.Append(' ');
            builder.Append("0 1");
        }

        return builder.Append("] >>").ToString();
    }

    private static string BuildInterpolationFunction(OfficeColor startColor, OfficeColor endColor) =>
        "<< /FunctionType 2 /Domain [0 1] /C0 [" +
        FormatColorComponent(startColor.R) + " " + FormatColorComponent(startColor.G) + " " + FormatColorComponent(startColor.B) +
        "] /C1 [" +
        FormatColorComponent(endColor.R) + " " + FormatColorComponent(endColor.G) + " " + FormatColorComponent(endColor.B) +
        "] /N 1 >>";

    private static void ValidateStops(IReadOnlyList<OfficeGradientStop>? stops) {
        if (stops == null || stops.Count < 2) throw new ArgumentException("An axial PDF shading needs at least two stops.", nameof(stops));
        if (!stops[0].Offset.Equals(0D) || !stops[stops.Count - 1].Offset.Equals(1D)) {
            throw new ArgumentException("Axial PDF shading stops must start at zero and end at one.", nameof(stops));
        }

        double previous = -1D;
        for (int index = 0; index < stops.Count; index++) {
            double offset = stops[index].Offset;
            if (double.IsNaN(offset) || double.IsInfinity(offset) || offset <= previous) {
                throw new ArgumentException("Axial PDF shading stops must use strictly increasing finite offsets.", nameof(stops));
            }

            previous = offset;
        }
    }

    private static string FormatColorComponent(byte value) =>
        FormatNumber(value / 255D);

    private static string FormatNumber(double value) =>
        value.ToString("0.###", CultureInfo.InvariantCulture);

    private static void ValidateOpacity(double value, string paramName) {
        ValidateFinite(value, paramName);
        if (value < 0 || value > 1) {
            throw new ArgumentOutOfRangeException(paramName, value, "PDF graphics-state opacity must be between 0 and 1.");
        }
    }

    private static void ValidateFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(paramName, value, "PDF visual resource numbers must be finite.");
        }
    }
}

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
        OfficeColor endColor) {
        ValidateFinite(x0, nameof(x0));
        ValidateFinite(y0, nameof(y0));
        ValidateFinite(x1, nameof(x1));
        ValidateFinite(y1, nameof(y1));

        return
            "<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [" +
            FormatNumber(x0) + " " + FormatNumber(y0) + " " + FormatNumber(x1) + " " + FormatNumber(y1) +
            "] /Function << /FunctionType 2 /Domain [0 1] /C0 [" +
            FormatColorComponent(startColor.R) + " " + FormatColorComponent(startColor.G) + " " + FormatColorComponent(startColor.B) +
            "] /C1 [" +
            FormatColorComponent(endColor.R) + " " + FormatColorComponent(endColor.G) + " " + FormatColorComponent(endColor.B) +
            "] /N 1 >> /Extend [true true] >>\n";
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

using OfficeIMO.Drawing;
using System.Globalization;

namespace OfficeIMO.Html;

internal static class HtmlCssRadialGradientParser {
    private const double MinimumRadius = 0.000001D;
    private static readonly double CornerScale = Math.Sqrt(2D);

    internal static bool TryParse(string? value, int maximumStops, out OfficeRadialGradient? gradient, out bool stopLimitExceeded) {
        gradient = null;
        stopLimitExceeded = false;
        if (string.IsNullOrWhiteSpace(value) || maximumStops < 2) return false;

        string text = value!.Trim();
        const string functionName = "radial-gradient";
        if (!text.StartsWith(functionName, StringComparison.OrdinalIgnoreCase)) return false;
        int open = functionName.Length;
        if (open >= text.Length || text[open] != '(' || text[text.Length - 1] != ')') return false;

        IReadOnlyList<string> arguments = HtmlRenderCssValues.SplitTopLevelCommas(text.Substring(open + 1, text.Length - open - 2));
        if (arguments.Count < 2) return false;
        int stopStart = 0;
        double centerX = 0.5D;
        double centerY = 0.5D;
        double radiusX = 0.5D * CornerScale;
        double radiusY = 0.5D * CornerScale;
        if (!HtmlCssGradientStops.IsColorStop(arguments[0])) {
            if (!TryParseEllipse(arguments[0], out centerX, out centerY, out radiusX, out radiusY)) return false;
            stopStart = 1;
        }

        if (!HtmlCssGradientStops.TryParse(arguments, stopStart, maximumStops, out IReadOnlyList<OfficeGradientStop>? stops, out stopLimitExceeded)
            || stops == null) {
            return false;
        }

        gradient = new OfficeRadialGradient(centerX, centerY, 0D, 0D, centerX, centerY, radiusX, radiusY, stops);
        return true;
    }

    private static bool TryParseEllipse(string value, out double centerX, out double centerY, out double radiusX, out double radiusY) {
        centerX = 0.5D;
        centerY = 0.5D;
        radiusX = 0.5D * CornerScale;
        radiusY = 0.5D * CornerScale;
        List<string> parts = HtmlRenderCssValues.SplitWhitespace(value)
            .Select(part => part.ToLowerInvariant())
            .ToList();
        int at = parts.IndexOf("at");
        if (at >= 0) {
            if (!TryParsePosition(parts.Skip(at + 1).ToList(), out centerX, out centerY)) return false;
            parts.RemoveRange(at, parts.Count - at);
        }

        if (parts.Remove("circle")) return false;
        parts.Remove("ellipse");
        if (parts.Count == 0) return ResolveExtent("farthest-corner", centerX, centerY, out radiusX, out radiusY);
        if (parts.Count == 1) return ResolveExtent(parts[0], centerX, centerY, out radiusX, out radiusY);
        if (parts.Count == 2
            && TryParseNonNegativePercentage(parts[0], out radiusX)
            && TryParseNonNegativePercentage(parts[1], out radiusY)) {
            radiusX = Math.Max(MinimumRadius, radiusX);
            radiusY = Math.Max(MinimumRadius, radiusY);
            return true;
        }

        return false;
    }

    private static bool ResolveExtent(string extent, double centerX, double centerY, out double radiusX, out double radiusY) {
        bool closest;
        bool corner;
        switch (extent) {
            case "closest-side":
                closest = true;
                corner = false;
                break;
            case "closest-corner":
                closest = true;
                corner = true;
                break;
            case "farthest-side":
                closest = false;
                corner = false;
                break;
            case "farthest-corner":
                closest = false;
                corner = true;
                break;
            default:
                radiusX = 0D;
                radiusY = 0D;
                return false;
        }

        double left = Math.Abs(centerX);
        double right = Math.Abs(1D - centerX);
        double top = Math.Abs(centerY);
        double bottom = Math.Abs(1D - centerY);
        radiusX = closest ? Math.Min(left, right) : Math.Max(left, right);
        radiusY = closest ? Math.Min(top, bottom) : Math.Max(top, bottom);
        if (corner) {
            radiusX *= CornerScale;
            radiusY *= CornerScale;
        }

        radiusX = Math.Max(MinimumRadius, radiusX);
        radiusY = Math.Max(MinimumRadius, radiusY);
        return true;
    }

    private static bool TryParsePosition(IReadOnlyList<string> parts, out double x, out double y) {
        x = 0.5D;
        y = 0.5D;
        if (parts.Count == 0 || parts.Count > 2) return false;
        if (parts.Count == 1) {
            if (TryParseHorizontalPosition(parts[0], out x)) return true;
            if (TryParseVerticalPosition(parts[0], out y)) return true;
            return false;
        }

        if (TryParseHorizontalPosition(parts[0], out x) && TryParseVerticalPosition(parts[1], out y)) return true;
        return TryParseHorizontalPosition(parts[1], out x) && TryParseVerticalPosition(parts[0], out y);
    }

    private static bool TryParseHorizontalPosition(string value, out double position) {
        switch (value) {
            case "left":
                position = 0D;
                return true;
            case "center":
                position = 0.5D;
                return true;
            case "right":
                position = 1D;
                return true;
            default:
                return TryParsePercentage(value, out position);
        }
    }

    private static bool TryParseVerticalPosition(string value, out double position) {
        switch (value) {
            case "top":
                position = 0D;
                return true;
            case "center":
                position = 0.5D;
                return true;
            case "bottom":
                position = 1D;
                return true;
            default:
                return TryParsePercentage(value, out position);
        }
    }

    private static bool TryParsePercentage(string value, out double result) {
        result = 0D;
        if (!value.EndsWith("%", StringComparison.Ordinal)
            || !double.TryParse(value.Substring(0, value.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out double percent)
            || double.IsNaN(percent)
            || double.IsInfinity(percent)) {
            return false;
        }

        result = percent / 100D;
        return true;
    }

    private static bool TryParseNonNegativePercentage(string value, out double result) =>
        TryParsePercentage(value, out result) && result >= 0D;

}

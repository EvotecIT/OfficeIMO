using OfficeIMO.Drawing;
using System.Globalization;

namespace OfficeIMO.Html;

internal static class HtmlCssLinearGradientParser {
    internal static bool TryParse(string? value, int maximumStops, out OfficeLinearGradient? gradient, out bool stopLimitExceeded) {
        gradient = null;
        stopLimitExceeded = false;
        if (string.IsNullOrWhiteSpace(value) || maximumStops < 2) return false;

        string text = value!.Trim();
        const string functionName = "linear-gradient";
        if (!text.StartsWith(functionName, StringComparison.OrdinalIgnoreCase)) return false;
        int open = functionName.Length;
        if (open >= text.Length || text[open] != '(' || text[text.Length - 1] != ')') return false;

        IReadOnlyList<string> arguments = HtmlRenderCssValues.SplitTopLevelCommas(text.Substring(open + 1, text.Length - open - 2));
        if (arguments.Count < 2) return false;

        double officeAngle = 90D;
        int stopStart = 0;
        if (TryParseDirection(arguments[0], out double parsedAngle)) {
            officeAngle = parsedAngle;
            stopStart = 1;
        }

        int stopCount = arguments.Count - stopStart;
        if (stopCount < 2) return false;
        if (stopCount > maximumStops) {
            stopLimitExceeded = true;
            return false;
        }
        var colors = new OfficeColor[stopCount];
        var offsets = new double?[stopCount];
        for (int index = 0; index < stopCount; index++) {
            if (!TryParseColorStop(arguments[index + stopStart], out colors[index], out offsets[index])) return false;
        }

        if (!TryResolveStops(colors, offsets, maximumStops, out IReadOnlyList<OfficeGradientStop>? stops, out stopLimitExceeded) || stops == null) return false;
        gradient = OfficeLinearGradient.FromAngle(stops, officeAngle);
        return true;
    }

    private static bool TryParseDirection(string value, out double officeAngle) {
        officeAngle = 0D;
        string normalized = value.Trim().ToLowerInvariant();
        if (normalized.StartsWith("to ", StringComparison.Ordinal)) {
            bool top = false;
            bool right = false;
            bool bottom = false;
            bool left = false;
            IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(normalized.Substring(3));
            if (parts.Count == 0 || parts.Count > 2) return false;
            foreach (string part in parts) {
                switch (part) {
                    case "top":
                        if (top || bottom) return false;
                        top = true;
                        break;
                    case "right":
                        if (right || left) return false;
                        right = true;
                        break;
                    case "bottom":
                        if (bottom || top) return false;
                        bottom = true;
                        break;
                    case "left":
                        if (left || right) return false;
                        left = true;
                        break;
                    default:
                        return false;
                }
            }

            if (top && right) officeAngle = 315D;
            else if (top && left) officeAngle = 225D;
            else if (bottom && right) officeAngle = 45D;
            else if (bottom && left) officeAngle = 135D;
            else if (top) officeAngle = 270D;
            else if (right) officeAngle = 0D;
            else if (bottom) officeAngle = 90D;
            else if (left) officeAngle = 180D;
            else return false;
            return true;
        }

        double multiplier;
        string number;
        if (normalized.EndsWith("deg", StringComparison.Ordinal)) {
            multiplier = 1D;
            number = normalized.Substring(0, normalized.Length - 3);
        } else if (normalized.EndsWith("grad", StringComparison.Ordinal)) {
            multiplier = 0.9D;
            number = normalized.Substring(0, normalized.Length - 4);
        } else if (normalized.EndsWith("rad", StringComparison.Ordinal)) {
            multiplier = 180D / Math.PI;
            number = normalized.Substring(0, normalized.Length - 3);
        } else if (normalized.EndsWith("turn", StringComparison.Ordinal)) {
            multiplier = 360D;
            number = normalized.Substring(0, normalized.Length - 4);
        } else {
            return false;
        }

        if (!double.TryParse(number.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double cssAngle)
            || double.IsNaN(cssAngle)
            || double.IsInfinity(cssAngle)) {
            return false;
        }

        officeAngle = NormalizeDegrees((cssAngle * multiplier) - 90D);
        return true;
    }

    private static bool TryParseColorStop(string value, out OfficeColor color, out double? offset) {
        color = default;
        offset = null;
        IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(value);
        if (parts.Count == 0 || parts.Count > 2 || !HtmlRenderCssValues.TryColor(parts[0], out color) || color.A != 255) return false;
        if (parts.Count == 1) return true;

        string position = parts[1].Trim();
        if (position == "0") {
            offset = 0D;
            return true;
        }

        if (!position.EndsWith("%", StringComparison.Ordinal)
            || !double.TryParse(position.Substring(0, position.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out double percent)
            || double.IsNaN(percent)
            || double.IsInfinity(percent)
            || percent < 0D
            || percent > 100D) {
            return false;
        }

        offset = percent / 100D;
        return true;
    }

    private static bool TryResolveStops(
        IReadOnlyList<OfficeColor> colors,
        double?[] offsets,
        int maximumStops,
        out IReadOnlyList<OfficeGradientStop>? stops,
        out bool stopLimitExceeded) {
        stops = null;
        stopLimitExceeded = false;
        offsets[0] ??= 0D;
        offsets[offsets.Length - 1] ??= 1D;

        int previousSpecified = 0;
        for (int index = 1; index < offsets.Length; index++) {
            if (!offsets[index].HasValue) continue;
            double previous = offsets[previousSpecified]!.Value;
            double current = offsets[index]!.Value;
            if (current <= previous) return false;
            int gap = index - previousSpecified;
            for (int fill = 1; fill < gap; fill++) {
                offsets[previousSpecified + fill] = previous + ((current - previous) * fill / gap);
            }

            previousSpecified = index;
        }

        bool addLeading = offsets[0]!.Value > 0D;
        bool addTrailing = offsets[offsets.Length - 1]!.Value < 1D;
        int resolvedCount = offsets.Length + (addLeading ? 1 : 0) + (addTrailing ? 1 : 0);
        if (resolvedCount > maximumStops) {
            stopLimitExceeded = true;
            return false;
        }

        var resolved = new List<OfficeGradientStop>(resolvedCount);
        if (addLeading) resolved.Add(new OfficeGradientStop(0D, colors[0]));
        for (int index = 0; index < offsets.Length; index++) {
            resolved.Add(new OfficeGradientStop(offsets[index]!.Value, colors[index]));
        }

        if (addTrailing) resolved.Add(new OfficeGradientStop(1D, colors[colors.Count - 1]));
        stops = resolved.AsReadOnly();
        return true;
    }

    private static double NormalizeDegrees(double value) {
        double normalized = value % 360D;
        return normalized < 0D ? normalized + 360D : normalized;
    }
}

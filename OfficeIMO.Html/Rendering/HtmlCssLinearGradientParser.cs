using OfficeIMO.Drawing;
using System.Globalization;

namespace OfficeIMO.Html;

internal static class HtmlCssLinearGradientParser {
    internal static bool TryParse(string? value, int maximumStops, out HtmlCssLinearGradientDefinition? definition, out bool stopLimitExceeded) {
        definition = null;
        stopLimitExceeded = false;
        if (string.IsNullOrWhiteSpace(value) || maximumStops < 2) return false;

        string text = value!.Trim();
        const string functionName = "linear-gradient";
        if (!text.StartsWith(functionName, StringComparison.OrdinalIgnoreCase)) return false;
        int open = functionName.Length;
        if (open >= text.Length || text[open] != '(' || text[text.Length - 1] != ')') return false;

        if (!HtmlRenderCssValues.TrySplitTopLevelCommas(
                text.Substring(open + 1, text.Length - open - 2),
                maximumStops == int.MaxValue ? int.MaxValue : maximumStops + 1,
                out IReadOnlyList<string> arguments)) {
            stopLimitExceeded = true;
            return false;
        }
        if (arguments.Count < 2) return false;

        double officeAngle = 90D;
        int stopStart = 0;
        if (TryParseDirection(arguments[0], out double parsedAngle)) {
            officeAngle = parsedAngle;
            stopStart = 1;
        }

        if (!HtmlCssGradientStops.TryParse(arguments, stopStart, maximumStops, out HtmlCssGradientStops? stops, out stopLimitExceeded) || stops == null) return false;
        definition = new HtmlCssLinearGradientDefinition(officeAngle, stops);
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

    private static double NormalizeDegrees(double value) {
        double normalized = value % 360D;
        return normalized < 0D ? normalized + 360D : normalized;
    }
}

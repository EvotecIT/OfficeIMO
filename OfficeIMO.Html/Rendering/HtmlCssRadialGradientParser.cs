using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal static class HtmlCssRadialGradientParser {
    private static readonly double CornerRadius = Math.Sqrt(0.5D);

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
        double radius = CornerRadius;
        if (!HtmlCssGradientStops.IsColorStop(arguments[0])) {
            if (!TryParseCenteredEllipse(arguments[0], out radius)) return false;
            stopStart = 1;
        }

        if (!HtmlCssGradientStops.TryParse(arguments, stopStart, maximumStops, out IReadOnlyList<OfficeGradientStop>? stops, out stopLimitExceeded)
            || stops == null) {
            return false;
        }

        gradient = new OfficeRadialGradient(0.5D, 0.5D, 0D, 0.5D, 0.5D, radius, stops);
        return true;
    }

    private static bool TryParseCenteredEllipse(string value, out double radius) {
        radius = CornerRadius;
        string normalized = string.Join(" ", HtmlRenderCssValues.SplitWhitespace(value).Select(part => part.ToLowerInvariant()));
        int at = normalized.IndexOf(" at ", StringComparison.Ordinal);
        string shapeAndSize = at >= 0 ? normalized.Substring(0, at).Trim() : normalized;
        string position = at >= 0 ? normalized.Substring(at + 4).Trim() : string.Empty;
        if (normalized.StartsWith("at ", StringComparison.Ordinal)) {
            shapeAndSize = string.Empty;
            position = normalized.Substring(3).Trim();
        }

        if (position.Length > 0 && position != "center" && position != "50% 50%") return false;
        if (shapeAndSize.Length == 0 || shapeAndSize == "ellipse" || shapeAndSize == "ellipse farthest-corner"
            || shapeAndSize == "farthest-corner ellipse" || shapeAndSize == "farthest-corner") {
            radius = CornerRadius;
            return true;
        }

        if (shapeAndSize == "ellipse closest-corner" || shapeAndSize == "closest-corner ellipse" || shapeAndSize == "closest-corner") {
            radius = CornerRadius;
            return true;
        }

        if (shapeAndSize == "ellipse closest-side" || shapeAndSize == "closest-side ellipse" || shapeAndSize == "closest-side"
            || shapeAndSize == "ellipse farthest-side" || shapeAndSize == "farthest-side ellipse" || shapeAndSize == "farthest-side") {
            radius = 0.5D;
            return true;
        }

        return false;
    }
}

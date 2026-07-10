using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal static class HtmlCssBoxShadowParser {
    internal static bool TryParse(
        string value,
        double fontSize,
        double rootFontSize,
        OfficeColor currentColor,
        out OfficeShadow? shadow) {
        shadow = null;
        string normalized = string.IsNullOrWhiteSpace(value) ? "none" : value.Trim().ToLowerInvariant();
        if (normalized == "none") return true;
        IReadOnlyList<string> layers = HtmlRenderCssValues.SplitTopLevelCommas(normalized);
        if (layers.Count != 1) return false;
        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(layers[0]);
        if (tokens.Count < 2 || tokens.Any(token => token == "inset")) return false;

        OfficeColor color = currentColor;
        bool colorSpecified = false;
        var lengths = new List<double>(4);
        foreach (string token in tokens) {
            if (string.Equals(token, "currentcolor", StringComparison.OrdinalIgnoreCase)) {
                if (colorSpecified) return false;
                color = currentColor;
                colorSpecified = true;
                continue;
            }
            if (HtmlRenderCssValues.TryColor(token, out OfficeColor parsedColor)) {
                if (colorSpecified) return false;
                color = parsedColor;
                colorSpecified = true;
                continue;
            }
            if (token.EndsWith("%", StringComparison.Ordinal)
                || !HtmlRenderCssValues.TryLength(token, 0D, fontSize, rootFontSize, out double length)) return false;
            lengths.Add(length);
        }

        if (lengths.Count < 2 || lengths.Count > 4) return false;
        double blur = lengths.Count > 2 ? lengths[2] : 0D;
        double spread = lengths.Count > 3 ? lengths[3] : 0D;
        if (blur < 0D || Math.Abs(spread) > 0.0001D) return false;
        double opacity = color.A / 255D;
        shadow = new OfficeShadow(OfficeColor.FromRgb(color.R, color.G, color.B), opacity, lengths[0], lengths[1], blur);
        return true;
    }

    internal static bool IsSupportedSyntax(string value) =>
        TryParse(value, 16D, 16D, OfficeColor.Black, out _);
}

using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal static class HtmlCssBoxShadowParser {
    internal static bool TryParse(
        string value,
        double fontSize,
        double rootFontSize,
        OfficeColor currentColor,
        out IReadOnlyList<HtmlCssBoxShadow> shadows) {
        shadows = Array.Empty<HtmlCssBoxShadow>();
        string normalized = string.IsNullOrWhiteSpace(value) ? "none" : value.Trim().ToLowerInvariant();
        if (normalized == "none") return true;
        IReadOnlyList<string> layers = HtmlRenderCssValues.SplitTopLevelCommas(normalized);
        if (layers.Count == 0) return false;
        var parsed = new List<HtmlCssBoxShadow>(layers.Count);
        for (int index = 0; index < layers.Count; index++) {
            if (!TryParseLayer(layers[index], fontSize, rootFontSize, currentColor, out HtmlCssBoxShadow? shadow)) return false;
            parsed.Add(shadow!);
        }
        shadows = parsed;
        return true;
    }

    private static bool TryParseLayer(
        string layer,
        double fontSize,
        double rootFontSize,
        OfficeColor currentColor,
        out HtmlCssBoxShadow? shadow) {
        shadow = null;
        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(layer);
        if (tokens.Count < 2) return false;

        OfficeColor color = currentColor;
        bool colorSpecified = false;
        bool inset = false;
        var lengths = new List<double>(4);
        foreach (string token in tokens) {
            if (token == "inset") {
                if (inset) return false;
                inset = true;
                continue;
            }
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
        if (blur < 0D) return false;
        double opacity = color.A / 255D;
        shadow = new HtmlCssBoxShadow(
            new OfficeShadow(OfficeColor.FromRgb(color.R, color.G, color.B), opacity, lengths[0], lengths[1], blur),
            spread,
            inset);
        return true;
    }

    internal static bool IsSupportedSyntax(string value) =>
        TryParse(value, 16D, 16D, OfficeColor.Black, out _);
}

internal sealed class HtmlCssBoxShadow {
    internal HtmlCssBoxShadow(OfficeShadow shadow, double spreadRadius, bool inset) {
        Shadow = shadow;
        SpreadRadius = spreadRadius;
        Inset = inset;
    }

    internal OfficeShadow Shadow { get; }
    internal double SpreadRadius { get; }
    internal bool Inset { get; }
}

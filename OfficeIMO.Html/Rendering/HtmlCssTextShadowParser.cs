using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal static class HtmlCssTextShadowParser {
    internal static bool TryParse(
        string value,
        double fontSize,
        double rootFontSize,
        double viewportWidth,
        double viewportHeight,
        double containerWidth,
        double containerHeight,
        OfficeColor currentColor,
        out IReadOnlyList<HtmlCssTextShadow> shadows) {
        shadows = Array.Empty<HtmlCssTextShadow>();
        string normalized = string.IsNullOrWhiteSpace(value) ? "none" : value.Trim().ToLowerInvariant();
        if (normalized == "none") return true;

        IReadOnlyList<string> layers = HtmlRenderCssValues.SplitTopLevelCommas(normalized);
        if (layers.Count == 0) return false;
        var parsed = new List<HtmlCssTextShadow>(layers.Count);
        foreach (string layer in layers) {
            if (!TryParseLayer(
                    layer,
                    fontSize,
                    rootFontSize,
                    viewportWidth,
                    viewportHeight,
                    containerWidth,
                    containerHeight,
                    currentColor,
                    out HtmlCssTextShadow? shadow)) return false;
            parsed.Add(shadow!);
        }

        shadows = parsed;
        return true;
    }

    private static bool TryParseLayer(
        string layer,
        double fontSize,
        double rootFontSize,
        double viewportWidth,
        double viewportHeight,
        double containerWidth,
        double containerHeight,
        OfficeColor currentColor,
        out HtmlCssTextShadow? shadow) {
        shadow = null;
        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(layer);
        if (tokens.Count < 2) return false;

        OfficeColor color = currentColor;
        bool colorSpecified = false;
        var lengths = new List<double>(3);
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
                || !HtmlRenderCssValues.TryLength(
                    token,
                    0D,
                    fontSize,
                    rootFontSize,
                    viewportWidth,
                    viewportHeight,
                    containerWidth,
                    containerHeight,
                    out double length)) return false;
            lengths.Add(length);
        }

        if (lengths.Count < 2 || lengths.Count > 3) return false;
        double blurRadius = lengths.Count == 3 ? lengths[2] : 0D;
        if (blurRadius < 0D) return false;
        shadow = new HtmlCssTextShadow(
            OfficeColor.FromRgb(color.R, color.G, color.B),
            color.A / 255D,
            lengths[0],
            lengths[1],
            blurRadius);
        return true;
    }

    internal static bool IsSupportedSyntax(string value) =>
        TryParse(value, 16D, 16D, 100D, 100D, 100D, 100D, OfficeColor.Black, out _);
}

internal sealed class HtmlCssTextShadow {
    internal HtmlCssTextShadow(OfficeColor color, double opacity, double offsetX, double offsetY, double blurRadius) {
        Color = color;
        Opacity = opacity;
        OffsetX = offsetX;
        OffsetY = offsetY;
        BlurRadius = blurRadius;
    }

    internal OfficeColor Color { get; }
    internal double Opacity { get; }
    internal double OffsetX { get; }
    internal double OffsetY { get; }
    internal double BlurRadius { get; }
}

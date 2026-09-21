using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal static class HtmlCssTextShadowParser {
    private const int MaximumParsedLayers = 64;
    internal static bool TryParse(
        string value,
        double fontSize,
        double rootFontSize,
        double viewportWidth,
        double viewportHeight,
        double containerWidth,
        double containerHeight,
        OfficeColor currentColor,
        out IReadOnlyList<HtmlCssTextShadow> shadows) =>
        TryParse(value, fontSize, rootFontSize, viewportWidth, viewportHeight,
            containerWidth, containerHeight, currentColor, 256, out shadows, out _);

    internal static bool TryParse(
        string value,
        double fontSize,
        double rootFontSize,
        double viewportWidth,
        double viewportHeight,
        double containerWidth,
        double containerHeight,
        OfficeColor currentColor,
        int maximumLayers,
        out IReadOnlyList<HtmlCssTextShadow> shadows,
        out int layerCount) {
        shadows = Array.Empty<HtmlCssTextShadow>();
        layerCount = 0;
        if (value == null || value.Length > 65536 || maximumLayers <= 0) return false;
        string normalized = string.IsNullOrWhiteSpace(value) ? "none" : value.Trim().ToLowerInvariant();
        if (normalized == "none") return true;

        var layers = new List<string>(Math.Min(maximumLayers, 16));
        int start = 0;
        int depth = 0;
        char quote = '\0';
        for (int index = 0; index < normalized.Length; index++) {
            char current = normalized[index];
            if (current == '\\' && index + 1 < normalized.Length) { index++; continue; }
            if (quote != '\0') { if (current == quote) quote = '\0'; continue; }
            if (current == '\'' || current == '"') quote = current;
            else if (current == '(') depth++;
            else if (current == ')' && depth > 0) depth--;
            else if (current == ',' && depth == 0) {
                if (layerCount >= MaximumParsedLayers) return false;
                layers.Add(normalized.Substring(start, index - start).Trim());
                layerCount++;
                start = index + 1;
            }
        }
        if (layerCount >= MaximumParsedLayers) return false;
        layers.Add(normalized.Substring(start).Trim());
        layerCount++;
        var parsed = new List<HtmlCssTextShadow>(Math.Min(layers.Count, maximumLayers));
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
            if (parsed.Count < maximumLayers) parsed.Add(shadow!);
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

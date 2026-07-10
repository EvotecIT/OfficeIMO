namespace OfficeIMO.Html;

internal static class HtmlCssTableParser {
    internal static bool TryParseBorderSpacing(string value, double fontSize, double rootFontSize, out double horizontal, out double vertical) {
        horizontal = 0D;
        vertical = 0D;
        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(value);
        if (tokens.Count == 0 || tokens.Count > 2) return false;
        if (!TryParseSpacingLength(tokens[0], fontSize, rootFontSize, out horizontal)) return false;
        vertical = horizontal;
        return tokens.Count == 1 || TryParseSpacingLength(tokens[1], fontSize, rootFontSize, out vertical);
    }

    private static bool TryParseSpacingLength(string value, double fontSize, double rootFontSize, out double result) {
        result = 0D;
        string normalized = value.Trim().ToLowerInvariant();
        return !normalized.EndsWith("%", StringComparison.Ordinal)
            && HtmlRenderCssValues.TryLength(normalized, 0D, fontSize, rootFontSize, out result)
            && result >= 0D;
    }
}

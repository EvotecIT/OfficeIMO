namespace OfficeIMO.Html;

internal static class HtmlCssTableParser {
    internal static bool TryParseBorderSpacing(string value, double fontSize, double rootFontSize, out double horizontal, out double vertical, double characterAdvance = double.NaN) =>
        TryParseBorderSpacing(value, fontSize, rootFontSize, 100D, 100D, out horizontal, out vertical, characterAdvance);

    internal static bool TryParseBorderSpacing(string value, double fontSize, double rootFontSize, double viewportWidth, double viewportHeight, out double horizontal, out double vertical, double characterAdvance = double.NaN) {
        horizontal = 0D;
        vertical = 0D;
        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(value);
        if (tokens.Count == 0 || tokens.Count > 2) return false;
        if (!TryParseSpacingLength(tokens[0], fontSize, rootFontSize, viewportWidth, viewportHeight, out double parsedHorizontal, characterAdvance)) return false;
        double parsedVertical = parsedHorizontal;
        if (tokens.Count == 2 && !TryParseSpacingLength(tokens[1], fontSize, rootFontSize, viewportWidth, viewportHeight, out parsedVertical, characterAdvance)) return false;
        horizontal = parsedHorizontal;
        vertical = parsedVertical;
        return true;
    }

    private static bool TryParseSpacingLength(string value, double fontSize, double rootFontSize, double viewportWidth, double viewportHeight, out double result, double characterAdvance = double.NaN) {
        result = 0D;
        string normalized = value.Trim().ToLowerInvariant();
        return !normalized.EndsWith("%", StringComparison.Ordinal)
            && HtmlRenderCssValues.TryLength(normalized, 0D, fontSize, rootFontSize, viewportWidth, viewportHeight, out result, characterAdvance)
            && result >= 0D;
    }
}

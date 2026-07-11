namespace OfficeIMO.Html;

internal static class HtmlCssOverflowClipMarginParser {
    internal static bool TryParse(
        string value,
        double fontSize,
        double rootFontSize,
        out string box,
        out double margin) {
        box = "padding-box";
        margin = 0D;
        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(value);
        if (tokens.Count == 0 || tokens.Count > 2) return false;

        bool boxSpecified = false;
        bool marginSpecified = false;
        foreach (string token in tokens) {
            string normalized = token.Trim().ToLowerInvariant();
            if (normalized == "content-box" || normalized == "padding-box" || normalized == "border-box") {
                if (boxSpecified) return false;
                box = normalized;
                boxSpecified = true;
                continue;
            }
            if (marginSpecified
                || normalized.EndsWith("%", StringComparison.Ordinal)
                || !HtmlRenderCssValues.TryLength(normalized, 0D, fontSize, rootFontSize, out double parsed)
                || parsed < 0D) return false;
            margin = parsed;
            marginSpecified = true;
        }
        return true;
    }
}

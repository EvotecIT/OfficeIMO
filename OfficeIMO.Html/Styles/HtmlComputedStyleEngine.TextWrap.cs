namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    /// <summary>Maps a legacy white-space keyword, or a collapse and wrap-mode pair, to its longhands.</summary>
    internal static bool TryExpandWhiteSpace(string value, out string collapse, out string wrapMode) {
        collapse = "collapse";
        wrapMode = "wrap";
        switch (value) {
            case "normal": return true;
            case "nowrap": wrapMode = "nowrap"; return true;
            case "pre": collapse = "preserve"; wrapMode = "nowrap"; return true;
            case "pre-wrap": collapse = "preserve"; return true;
            case "pre-line": collapse = "preserve-breaks"; return true;
            case "break-spaces": collapse = "break-spaces"; return true;
        }
        bool hasCollapse = false, hasWrap = false;
        IReadOnlyList<string> keywords = HtmlRenderCssValues.SplitWhitespace(value);
        if (keywords.Count == 0) return false;
        foreach (string keyword in keywords) {
            if (keyword is "collapse" or "preserve" or "preserve-breaks" or "preserve-spaces" or "break-spaces" && !hasCollapse) {
                collapse = keyword;
                hasCollapse = true;
            } else if (keyword is "wrap" or "nowrap" && !hasWrap) {
                wrapMode = keyword;
                hasWrap = true;
            } else {
                return false;
            }
        }
        return true;
    }

    /// <summary>Maps text-wrap: [ wrap | nowrap ] || [ auto | balance | stable | pretty ] to its longhands.</summary>
    internal static bool TryExpandTextWrap(string value, out string mode, out string style) {
        mode = "wrap";
        style = "auto";
        bool hasMode = false, hasStyle = false;
        IReadOnlyList<string> keywords = HtmlRenderCssValues.SplitWhitespace(value);
        if (keywords.Count == 0) return false;
        foreach (string keyword in keywords) {
            if (keyword is "wrap" or "nowrap" && !hasMode) {
                mode = keyword;
                hasMode = true;
            } else if (keyword is "auto" or "balance" or "stable" or "pretty" && !hasStyle) {
                style = keyword;
                hasStyle = true;
            } else {
                return false;
            }
        }
        return true;
    }
}

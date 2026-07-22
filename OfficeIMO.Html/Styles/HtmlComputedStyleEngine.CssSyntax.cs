namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private static bool IsInsideCssComment(string text, int index) {
        int open = text.LastIndexOf("/*", Math.Max(0, index), StringComparison.Ordinal);
        if (open < 0) {
            return false;
        }

        int close = text.LastIndexOf("*/", Math.Max(0, index), StringComparison.Ordinal);
        return close < open;
    }

    private static string StripCssCommentsOutsideStrings(string css) {
        var result = new System.Text.StringBuilder(css.Length);
        char quote = '\0';
        for (int i = 0; i < css.Length; i++) {
            char current = css[i];
            if (quote != '\0') {
                result.Append(current);
                if (current == quote && !IsEscaped(css, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
                result.Append(current);
                continue;
            }

            if (current == '/' && i + 1 < css.Length && css[i + 1] == '*') {
                i += 2;
                while (i + 1 < css.Length && !(css[i] == '*' && css[i + 1] == '/')) {
                    i++;
                }

                if (i + 1 < css.Length) {
                    i++;
                }

                result.Append(' ');
                continue;
            }

            result.Append(current);
        }

        return result.ToString();
    }

    internal static IEnumerable<string> SplitSelectorList(string selectorText) {
        int depth = 0;
        char quote = '\0';
        int start = 0;
        for (int i = 0; i < selectorText.Length; i++) {
            char current = selectorText[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(selectorText, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
                continue;
            }

            if (current == '(' || current == '[') {
                depth++;
                continue;
            }

            if (current == ')' || current == ']') {
                if (depth > 0) {
                    depth--;
                }

                continue;
            }

            if (current == ',' && depth == 0) {
                yield return selectorText.Substring(start, i - start);
                start = i + 1;
            }
        }

        yield return selectorText.Substring(start);
    }

    private static IEnumerable<string> SplitCssDeclarations(string styleText) {
        int depth = 0;
        char quote = '\0';
        int start = 0;
        for (int i = 0; i < styleText.Length; i++) {
            char current = styleText[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(styleText, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
                continue;
            }

            if (current == '(') {
                depth++;
                continue;
            }

            if (current == ')') {
                if (depth > 0) {
                    depth--;
                }

                continue;
            }

            if (current == ';' && depth == 0) {
                yield return styleText.Substring(start, i - start);
                start = i + 1;
            }
        }

        yield return styleText.Substring(start);
    }

    private static bool IsInsideCssString(string text, int index) {
        char quote = '\0';
        for (int i = 0; i < index && i < text.Length; i++) {
            char current = text[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(text, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
            }
        }

        return quote != '\0';
    }

    private static bool IsEscaped(string text, int index) {
        int slashCount = 0;
        for (int i = index - 1; i >= 0 && text[i] == '\\'; i--) {
            slashCount++;
        }

        return slashCount % 2 == 1;
    }

}

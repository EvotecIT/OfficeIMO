namespace OfficeIMO.Html;

internal static class HtmlRawTextScanner {
    internal static int FindClosingTag(string html, int offset, string tagName) {
        if (tagName.Equals("script", StringComparison.OrdinalIgnoreCase)) return FindScriptClosingTag(html, offset);
        string closingPrefix = "</" + tagName;
        int candidate = offset;
        while (candidate < html.Length) {
            candidate = html.IndexOf(closingPrefix, candidate, StringComparison.OrdinalIgnoreCase);
            if (candidate < 0) return -1;
            int delimiter = candidate + closingPrefix.Length;
            if (delimiter >= html.Length || IsAsciiWhitespace(html[delimiter]) || html[delimiter] is '>' or '/') return candidate;
            candidate = delimiter;
        }
        return -1;
    }

    private static int FindScriptClosingTag(string html, int offset) {
        ScriptTextState state = ScriptTextState.Normal;
        for (int index = offset; index < html.Length; index++) {
            if (state != ScriptTextState.DoubleEscaped && MatchesScriptTag(html, index, "</script")) return index;
            if (state == ScriptTextState.Normal && StartsWith(html, index, "<!--", StringComparison.Ordinal)) {
                state = ScriptTextState.Escaped;
                index += 3;
                continue;
            }
            if (state != ScriptTextState.Normal && StartsWith(html, index, "-->", StringComparison.Ordinal)) {
                state = ScriptTextState.Normal;
                index += 2;
                continue;
            }
            if (state == ScriptTextState.Escaped && MatchesScriptTag(html, index, "<script")) {
                state = ScriptTextState.DoubleEscaped;
                index += 6;
                continue;
            }
            if (state == ScriptTextState.DoubleEscaped && MatchesScriptTag(html, index, "</script")) {
                state = ScriptTextState.Escaped;
                index += 7;
            }
        }
        return -1;
    }

    private static bool MatchesScriptTag(string html, int offset, string prefix) {
        if (!StartsWith(html, offset, prefix, StringComparison.OrdinalIgnoreCase)) return false;
        int delimiter = offset + prefix.Length;
        return delimiter >= html.Length || IsAsciiWhitespace(html[delimiter]) || html[delimiter] is '>' or '/';
    }

    private static bool StartsWith(string value, int offset, string prefix, StringComparison comparison) =>
        offset <= value.Length - prefix.Length && string.Compare(value, offset, prefix, 0, prefix.Length, comparison) == 0;

    private static bool IsAsciiWhitespace(char value) => value is '\t' or '\n' or '\f' or '\r' or ' ';

    private enum ScriptTextState { Normal, Escaped, DoubleEscaped }
}

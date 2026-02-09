namespace OfficeIMO.MarkdownRenderer;

internal static class JavaScriptString {
    /// <summary>
    /// Encodes a string as a JavaScript string literal (single-quoted).
    /// Intended for small snippets passed into WebView2.ExecuteScriptAsync(...).
    /// </summary>
    internal static string SingleQuoted(string? value) {
        if (value == null) return "''";
        var sb = new StringBuilder(value.Length + 16);
        sb.Append('\'');
        for (int i = 0; i < value.Length; i++) {
            char c = value[i];
            switch (c) {
                case '\\': sb.Append(@"\\"); break;
                case '\'': sb.Append(@"\'"); break;
                case '\r': sb.Append(@"\r"); break;
                case '\n': sb.Append(@"\n"); break;
                case '\t': sb.Append(@"\t"); break;
                case '\u2028': sb.Append(@"\u2028"); break; // JS line separator
                case '\u2029': sb.Append(@"\u2029"); break; // JS paragraph separator
                default:
                    if (c < 32) {
                        sb.Append(@"\x");
                        sb.Append(((int)c).ToString("x2"));
                    } else {
                        sb.Append(c);
                    }
                    break;
            }
        }
        sb.Append('\'');
        return sb.ToString();
    }
}


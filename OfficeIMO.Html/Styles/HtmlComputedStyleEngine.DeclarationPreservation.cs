namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private static string RestoreManagedDeclarationName(string name) {
        string? prefix = name.StartsWith(FontDeclarationSentinelPrefix, StringComparison.OrdinalIgnoreCase)
            ? FontDeclarationSentinelPrefix
            : name.StartsWith(LayoutDeclarationSentinelPrefix, StringComparison.OrdinalIgnoreCase) ? LayoutDeclarationSentinelPrefix : null;
        if (prefix == null) return name;
        string suffix = name.Substring(prefix.Length);
        int separator = suffix.IndexOf('-');
        return separator >= 0 ? suffix.Substring(separator + 1) : name;
    }

    // Keep authored order and values where parser expansion loses font syntax,
    // subgrid tracks, row/column gap order, or grid placement shorthands.
    private static string PreserveManagedDeclarations(string css) {
        var result = new System.Text.StringBuilder(css.Length);
        int copied = 0;
        int declarationId = 0;
        for (int index = 0; index < css.Length; index++) {
            char current = css[index];
            if (current is '\'' or '"') {
                char quote = current;
                while (++index < css.Length && (css[index] != quote || IsEscaped(css, index))) { }
                continue;
            }
            if (current == '/' && index + 1 < css.Length && css[index + 1] == '*') {
                int end = css.IndexOf("*/", index + 2, StringComparison.Ordinal);
                index = end < 0 ? css.Length : end + 1;
                continue;
            }
            int before = SkipCssWhitespaceAndCommentsBackward(css, index - 1);
            if (before < 0 || css[before] is not ('{' or ';')) continue;
            int endName = index;
            if (!HtmlCssIdentifierParser.TryRead(css, ref endName, out string name)) continue;
            name = name.ToLowerInvariant();
            string? prefix = name == "font" || FontShorthandLonghands.Contains(name)
                ? FontDeclarationSentinelPrefix : LayoutDeclarationNames.Contains(name) ? LayoutDeclarationSentinelPrefix : null;
            if (prefix == null) continue;
            int colon = SkipCssWhitespaceAndCommentsForward(css, endName);
            if (colon >= css.Length || css[colon] != ':') continue;
            result.Append(css, copied, index - copied).Append(prefix).Append(declarationId++).Append('-').Append(name);
            copied = endName;
            index = FindDeclarationValueEnd(css, colon + 1) - 1;
        }
        if (copied == 0) return css;
        result.Append(css, copied, css.Length - copied);
        return result.ToString();
    }
}

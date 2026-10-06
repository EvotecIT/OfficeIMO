namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private const string BorderDeclarationSentinelPrefix = "-officeimo-internal-border-declaration-";
    private static readonly HashSet<string> BorderDeclarationNames = new(StringComparer.Ordinal) {
        "border", "border-width", "border-style", "border-color",
        "border-top", "border-right", "border-bottom", "border-left",
        "border-top-width", "border-right-width", "border-bottom-width", "border-left-width",
        "border-top-style", "border-right-style", "border-bottom-style", "border-left-style",
        "border-top-color", "border-right-color", "border-bottom-color", "border-left-color"
    };

    private static string RestoreManagedDeclarationName(string name) {
        string? prefix = name.StartsWith(FontDeclarationSentinelPrefix, StringComparison.OrdinalIgnoreCase)
            ? FontDeclarationSentinelPrefix
            : name.StartsWith(LayoutDeclarationSentinelPrefix, StringComparison.OrdinalIgnoreCase) ? LayoutDeclarationSentinelPrefix
            : name.StartsWith(BorderDeclarationSentinelPrefix, StringComparison.OrdinalIgnoreCase) ? BorderDeclarationSentinelPrefix : null;
        if (prefix == null) return name;
        string suffix = name.Substring(prefix.Length);
        int separator = suffix.IndexOf('-');
        return separator >= 0 ? suffix.Substring(separator + 1) : name;
    }

    // Keep authored order and values where parser expansion loses font syntax,
    // subgrid tracks, row/column gap order, grid placement shorthands, or
    // variable-backed border components. A border-color variable must not become
    // a synthesized colour-only border shorthand that resets width and style.
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
            bool borderDeclaration = BorderDeclarationNames.Contains(name);
            if (prefix == null && !borderDeclaration) continue;
            int colon = SkipCssWhitespaceAndCommentsForward(css, endName);
            if (colon >= css.Length || css[colon] != ':') continue;
            int valueEnd = FindDeclarationValueEnd(css, colon + 1);
            if (prefix == null) {
                if (!HtmlCssCustomPropertyResolver.ContainsVarFunction(css.Substring(colon + 1, valueEnd - colon - 1))) continue;
                prefix = BorderDeclarationSentinelPrefix;
            }
            result.Append(css, copied, index - copied).Append(prefix).Append(declarationId++).Append('-').Append(name);
            copied = endName;
            index = valueEnd - 1;
        }
        if (copied == 0) return css;
        result.Append(css, copied, css.Length - copied);
        return result.ToString();
    }
}

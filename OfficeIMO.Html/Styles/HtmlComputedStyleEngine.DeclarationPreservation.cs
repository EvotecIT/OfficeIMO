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
    // intrinsic dimensions, subgrid tracks, row/column gap order, grid placement shorthands, or
    // border and text-decoration components. Preserve the whole family: mixing preserved variables
    // with parser-collapsed ordinary declarations loses their relative order.
    // A colour-only override must not reset the authored width and style.
    private static string PreserveManagedDeclarations(string css, OfficeIMO.Html.Css.HtmlCssStyleSheet sheet) {
        var result = new System.Text.StringBuilder(css.Length);
        int copied = 0;
        int declarationId = 0;
        // The owned parser already distinguishes declaration names from URL,
        // function and custom-property component values. Reuse its source spans
        // rather than treating every semicolon or brace as a declaration boundary.
        foreach (OfficeIMO.Html.Css.HtmlCssDeclaration declaration in EnumerateManagedDeclarationCandidates(sheet.Rules)) {
            int index = declaration.Span.Offset;
            int endName = index;
            if (!HtmlCssIdentifierParser.TryRead(css, ref endName, out string name)) continue;
            name = name.ToLowerInvariant();
            string? prefix = name == "font" || FontShorthandLonghands.Contains(name)
                ? FontDeclarationSentinelPrefix
                : LayoutDeclarationNames.Contains(name) || DimensionDeclarationNames.Contains(name) ? LayoutDeclarationSentinelPrefix
                : BorderDeclarationNames.Contains(name) ? BorderDeclarationSentinelPrefix : null;
            if (prefix == null) continue;
            result.Append(css, copied, index - copied).Append(prefix).Append(declarationId++).Append('-').Append(name);
            copied = endName;
        }
        if (copied == 0) return css;
        result.Append(css, copied, css.Length - copied);
        return result.ToString();
    }

    private static IEnumerable<OfficeIMO.Html.Css.HtmlCssDeclaration> EnumerateManagedDeclarationCandidates(
        IEnumerable<OfficeIMO.Html.Css.HtmlCssSyntaxNode> nodes) {
        foreach (OfficeIMO.Html.Css.HtmlCssSyntaxNode node in nodes) {
            if (node is OfficeIMO.Html.Css.HtmlCssDeclaration declaration) yield return declaration;
            else if (node is OfficeIMO.Html.Css.HtmlCssRule rule) {
                foreach (OfficeIMO.Html.Css.HtmlCssDeclaration child in EnumerateManagedDeclarationCandidates(rule.Contents)) {
                    yield return child;
                }
            }
        }
    }
}

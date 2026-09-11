using AngleSharp.Css.Parser;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private const string FontDeclarationSentinelPrefix = "-officeimo-internal-font-declaration-";
    private static readonly string[] FontShorthandLonghands = {
        "font-style", "font-variant", "font-weight", "font-stretch", "font-size", "line-height", "font-family",
        "font-feature-settings", "font-kerning", "font-variant-caps", "font-variant-east-asian",
        "font-variant-ligatures", "font-variant-numeric"
    };

    private static bool TryExpandFontShorthand(string value, out IReadOnlyList<KeyValuePair<string, string>> longhands) {
        if (IsCssWideKeyword(value.Trim())) {
            longhands = FontShorthandLonghands.Select(name => new KeyValuePair<string, string>(name, value)).ToArray();
            return true;
        }
        // Use the same CSS parser as stylesheet rules, retaining the declaration's
        // original cascade priority and order when applying its longhands.
        var parsed = new CssParser().ParseDeclaration("font:" + value);
        var expanded = new List<KeyValuePair<string, string>>();
        if (parsed != null) {
            foreach (string name in FontShorthandLonghands) {
                string resolved = parsed.GetPropertyValue(name);
                if (!string.IsNullOrWhiteSpace(resolved)) expanded.Add(new KeyValuePair<string, string>(name, resolved));
            }
        }
        if (!expanded.Any(item => item.Key == "font-size") || !expanded.Any(item => item.Key == "font-family")) {
            longhands = Array.Empty<KeyValuePair<string, string>>();
            return false;
        }
        foreach (string name in FontShorthandLonghands) {
            if (!expanded.Any(item => item.Key == name)) {
                // Reset-only font subproperties participate at the shorthand's
                // cascade position. Initial also prevents inherited settings
                // from surviving this declaration; font-palette is independent.
                bool resetOnly = name == "font-feature-settings" || name == "font-kerning"
                    || name.StartsWith("font-variant-", StringComparison.Ordinal);
                expanded.Add(new KeyValuePair<string, string>(name, resetOnly ? "initial" : "normal"));
            }
        }
        longhands = expanded;
        return true;
    }

    private static void ResolveDeferredFontLonghands(Dictionary<string, string> properties, ISet<string> deferred,
        IReadOnlyDictionary<string, string>? parentProperties, ISet<string> inherited, ISet<string> reset) {
        foreach (string name in deferred) {
            string value = "unset";
            if (HtmlCssCustomPropertyResolver.TryResolve(properties[name],
                    customName => properties.TryGetValue(customName, out string? local) ? local
                        : parentProperties != null && parentProperties.TryGetValue(customName, out string? parent) ? parent : null,
                    out string shorthand)
                && TryExpandFontShorthand(shorthand, out IReadOnlyList<KeyValuePair<string, string>> longhands)) {
                value = longhands.First(item => item.Key == name).Value;
            }
            // Invalid substitution still occupies the shorthand's cascade position.
            // Resolve its longhands as unset rather than reviving earlier declarations.
            CssKeywordResolution resolved = ResolveCssWideKeyword(name, value, parentProperties);
            if (resolved.HasValue) {
                properties[name] = resolved.Value;
                reset.Remove(name);
                if (resolved.InheritsComputedValue) inherited.Add(name); else inherited.Remove(name);
            } else {
                properties.Remove(name);
                inherited.Remove(name);
                reset.Add(name);
            }
        }
    }

    private static string RestoreFontShorthandName(string name) {
        if (!name.StartsWith(FontDeclarationSentinelPrefix, StringComparison.OrdinalIgnoreCase)) return name;
        string suffix = name.Substring(FontDeclarationSentinelPrefix.Length);
        int separator = suffix.IndexOf('-');
        return separator >= 0 ? suffix.Substring(separator + 1) : name;
    }

    private static string PreserveFontShorthandDeclarations(string css) {
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
            if (name != "font" && !FontShorthandLonghands.Contains(name)) continue;
            int colon = SkipCssWhitespaceAndCommentsForward(css, endName);
            if (colon >= css.Length || css[colon] != ':') continue;
            result.Append(css, copied, index - copied).Append(FontDeclarationSentinelPrefix).Append(declarationId++).Append('-').Append(name);
            copied = endName;
            index = FindDeclarationValueEnd(css, colon + 1) - 1;
        }
        if (copied == 0) return css;
        result.Append(css, copied, css.Length - copied);
        return result.ToString();
    }
}

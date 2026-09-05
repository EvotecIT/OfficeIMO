using AngleSharp.Css.Dom;
using AngleSharp.Css.Parser;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

internal sealed class HtmlCssFontFaceDefinition {
    internal HtmlCssFontFaceDefinition(string familyName, string source, string weight, string stretch, string style, string unicodeRange) {
        FamilyName = CleanFamilyName(familyName);
        Source = source ?? string.Empty;
        Weight = weight ?? string.Empty;
        Stretch = stretch ?? string.Empty;
        Style = style ?? string.Empty;
        UnicodeRange = unicodeRange ?? string.Empty;
    }

    internal string FamilyName { get; }
    internal string Source { get; }
    internal string Weight { get; }
    internal string Stretch { get; }
    internal string Style { get; }
    internal string UnicodeRange { get; }

    internal HtmlCssFontFaceDefinition WithRawDescriptors(IReadOnlyDictionary<string, string> declarations) =>
        new HtmlCssFontFaceDefinition(
            declarations.TryGetValue("font-family", out string? family) ? family : FamilyName,
            declarations.TryGetValue("src", out string? source) ? source : Source,
            declarations.TryGetValue("font-weight", out string? weight) ? weight : Weight,
            declarations.TryGetValue("font-stretch", out string? stretch) ? stretch : Stretch,
            declarations.TryGetValue("font-style", out string? style) ? style : Style,
            declarations.TryGetValue("unicode-range", out string? range) ? range : UnicodeRange);

    private static string CleanFamilyName(string value) {
        string family = (value ?? string.Empty).Trim();
        while (family.Length >= 2
               && ((family[0] == '"' && family[family.Length - 1] == '"')
                   || (family[0] == '\'' && family[family.Length - 1] == '\''))) {
            family = family.Substring(1, family.Length - 2).Trim();
        }

        return family;
    }
}

public static partial class HtmlResourcePipeline {
    internal static IReadOnlyList<HtmlCssFontFaceDefinition> ExtractFontFaces(string css, HtmlResourcePipelineOptions options) {
        var definitions = new List<HtmlCssFontFaceDefinition>();
        if (string.IsNullOrWhiteSpace(css)) {
            return definitions.AsReadOnly();
        }

        HtmlCssRuleBlockScanner.ValidateStylesheet(css, options.Limits);
        var parser = new CssParser();
        ICssStyleSheet stylesheet = parser.ParseStyleSheet(css);
        IReadOnlyList<IReadOnlyDictionary<string, string>> rawRules = ExtractRawFontFaceDescriptors(css);
        var rawDescriptors = new Dictionary<ICssFontFaceRule, IReadOnlyDictionary<string, string>>();
        int rawRuleIndex = 0;
        foreach (ICssRule rule in stylesheet.Rules) {
            MapRawFontFaceDescriptors(rule, rawRules, rawDescriptors, ref rawRuleIndex);
        }
        foreach (ICssRule rule in stylesheet.Rules) {
            AddFontFaces(rule, options, definitions, rawDescriptors);
        }

        return definitions.AsReadOnly();
    }

    internal static IReadOnlyList<string> ExtractFontFaceUrls(string source) {
        return ExtractCssUrls(source);
    }

    internal static IReadOnlyList<string> ExtractCssUrls(string source) {
        var urls = new List<string>();
        if (string.IsNullOrWhiteSpace(source)) {
            return urls.AsReadOnly();
        }

        foreach (Match match in CssUrlExpression.Matches(source)) {
            if (!IsValidCssUrlMatch(source, match) || !IsCssFunctionNameAt(source, match.Index, "url") || IsInsideCssString(source, match.Index)) {
                continue;
            }

            string value = DecodeCssEscapes(match.Groups["url"].Value.Trim().Trim('\'', '"'));
            if (!string.IsNullOrWhiteSpace(value) && !IsFragmentOnlyReference(value)) {
                urls.Add(value);
            }
        }

        return urls.AsReadOnly();
    }

    internal static string RebaseExternalStylesheetUrls(string css, Uri baseUri, HtmlUrlPolicy policy) {
        if (string.IsNullOrWhiteSpace(css)) {
            return css ?? string.Empty;
        }

        var replacements = new List<(int Start, int Length, string Value)>();
        HtmlUrlPolicy resourcePolicy = HtmlResourceUrlPolicy.Create(policy);
        foreach (Match match in CssUrlExpression.Matches(css)) {
            if (!IsValidCssUrlMatch(css, match) || !IsCssFunctionNameAt(css, match.Index, "url") || IsInsideCssString(css, match.Index)) {
                continue;
            }

            string source = DecodeCssEscapes(match.Groups["url"].Value.Trim().Trim('\'', '"'));
            if (string.IsNullOrWhiteSpace(source) || IsFragmentOnlyReference(source)) {
                continue;
            }

            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(source, baseUri, resourcePolicy);
            string replacement = resolved.Length == 0
                ? "url(\"\")"
                : "url(\"" + EscapeCssString(resolved) + "\")";
            replacements.Add((match.Index, match.Length, replacement));
        }

        if (replacements.Count == 0) {
            return css;
        }

        var builder = new StringBuilder(css);
        for (int index = replacements.Count - 1; index >= 0; index--) {
            (int start, int length, string value) = replacements[index];
            builder.Remove(start, length);
            builder.Insert(start, value);
        }

        return builder.ToString();
    }

    private static void AddFontFaces(
        ICssRule rule,
        HtmlResourcePipelineOptions options,
        ICollection<HtmlCssFontFaceDefinition> definitions,
        IReadOnlyDictionary<ICssFontFaceRule, IReadOnlyDictionary<string, string>> rawDescriptors) {
        if (rule is ICssMediaRule mediaRule && !IsApplicableMedia(mediaRule.ConditionText, options)) {
            return;
        }

        if (rule is ICssSupportsRule supportsRule && !HtmlComputedStyleEngine.IsApplicableSupports(supportsRule.ConditionText)) {
            return;
        }

        if (rule is ICssFontFaceRule fontFace) {
            var definition = new HtmlCssFontFaceDefinition(
                fontFace.Family,
                fontFace.Source,
                fontFace.Weight,
                fontFace.Stretch,
                fontFace.Style,
                fontFace.Range);
            definitions.Add(rawDescriptors.TryGetValue(fontFace, out IReadOnlyDictionary<string, string>? raw)
                ? definition.WithRawDescriptors(raw)
                : definition);
            return;
        }

        if (rule is ICssGroupingRule groupingRule) {
            foreach (ICssRule child in groupingRule.Rules) {
                AddFontFaces(child, options, definitions, rawDescriptors);
            }
        }
    }

    private static IReadOnlyList<IReadOnlyDictionary<string, string>> ExtractRawFontFaceDescriptors(string css) {
        IReadOnlyDictionary<int, int> closures = HtmlCssRuleBlockScanner.Scan(css, new HtmlCssProcessingBudget(null));
        var rawRules = new List<IReadOnlyDictionary<string, string>>();
        foreach (KeyValuePair<int, int> closure in closures.OrderBy(item => item.Key)) {
            int cursor = closure.Key - 1;
            while (cursor >= 0 && char.IsWhiteSpace(css[cursor])) cursor--;
            const string RuleName = "@font-face";
            int ruleStart = cursor - RuleName.Length + 1;
            if (ruleStart < 0
                || !css.Substring(ruleStart, RuleName.Length).Equals(RuleName, StringComparison.OrdinalIgnoreCase)
                || ruleStart > 0 && (char.IsLetterOrDigit(css[ruleStart - 1]) || css[ruleStart - 1] == '-' || css[ruleStart - 1] == '_')) {
                continue;
            }
            string body = css.Substring(closure.Key + 1, closure.Value - closure.Key - 1);
            rawRules.Add(HtmlRenderCssValues.ParseInlineStyleDeclarations(body));
        }
        return rawRules.AsReadOnly();
    }

    private static void MapRawFontFaceDescriptors(
        ICssRule rule,
        IReadOnlyList<IReadOnlyDictionary<string, string>> rawRules,
        IDictionary<ICssFontFaceRule, IReadOnlyDictionary<string, string>> mapped,
        ref int rawRuleIndex) {
        if (rule is ICssFontFaceRule fontFace) {
            while (rawRuleIndex < rawRules.Count) {
                IReadOnlyDictionary<string, string> raw = rawRules[rawRuleIndex++];
                if (!raw.TryGetValue("font-family", out string? rawFamily)
                    || !string.Equals(CleanFontFamily(rawFamily), CleanFontFamily(fontFace.Family), StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }
                mapped[fontFace] = raw;
                break;
            }
            return;
        }
        if (rule is ICssGroupingRule groupingRule) {
            foreach (ICssRule child in groupingRule.Rules) {
                MapRawFontFaceDescriptors(child, rawRules, mapped, ref rawRuleIndex);
            }
        }
    }

    private static string CleanFontFamily(string value) {
        string family = (value ?? string.Empty).Trim();
        while (family.Length >= 2
               && ((family[0] == '"' && family[family.Length - 1] == '"')
                   || (family[0] == '\'' && family[family.Length - 1] == '\''))) {
            family = family.Substring(1, family.Length - 2).Trim();
        }
        return family;
    }

    private static string EscapeCssString(string value) {
        var builder = new StringBuilder(value.Length);
        foreach (char character in value) {
            if (character == '\\' || character == '"' || character == '\r' || character == '\n' || character == '\f') {
                builder.Append('\\').Append(character);
            } else {
                builder.Append(character);
            }
        }

        return builder.ToString();
    }
}

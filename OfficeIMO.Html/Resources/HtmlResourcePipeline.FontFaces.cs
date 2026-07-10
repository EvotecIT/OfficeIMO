using AngleSharp.Css.Dom;
using AngleSharp.Css.Parser;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

internal sealed class HtmlCssFontFaceDefinition {
    internal HtmlCssFontFaceDefinition(string familyName, string source, string weight, string style) {
        FamilyName = CleanFamilyName(familyName);
        Source = source ?? string.Empty;
        Weight = weight ?? string.Empty;
        Style = style ?? string.Empty;
    }

    internal string FamilyName { get; }
    internal string Source { get; }
    internal string Weight { get; }
    internal string Style { get; }

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
    internal static IReadOnlyList<HtmlCssFontFaceDefinition> ExtractFontFaces(string css, HtmlCssMediaContext mediaContext) {
        var definitions = new List<HtmlCssFontFaceDefinition>();
        if (string.IsNullOrWhiteSpace(css)) {
            return definitions.AsReadOnly();
        }

        var parser = new CssParser();
        ICssStyleSheet stylesheet = parser.ParseStyleSheet(css);
        foreach (ICssRule rule in stylesheet.Rules) {
            AddFontFaces(rule, mediaContext, definitions);
        }

        return definitions.AsReadOnly();
    }

    internal static IReadOnlyList<string> ExtractFontFaceUrls(string source) {
        var urls = new List<string>();
        if (string.IsNullOrWhiteSpace(source)) {
            return urls.AsReadOnly();
        }

        foreach (Match match in CssUrlExpression.Matches(source)) {
            if (!IsCssFunctionNameAt(source, match.Index, "url") || IsInsideCssString(source, match.Index)) {
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
            if (!IsCssFunctionNameAt(css, match.Index, "url") || IsInsideCssString(css, match.Index)) {
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

    private static void AddFontFaces(ICssRule rule, HtmlCssMediaContext mediaContext, ICollection<HtmlCssFontFaceDefinition> definitions) {
        if (rule is ICssMediaRule mediaRule && !HtmlComputedStyleEngine.IsApplicableMedia(mediaRule.ConditionText, mediaContext)) {
            return;
        }

        if (rule is ICssSupportsRule supportsRule && !HtmlComputedStyleEngine.IsApplicableSupports(supportsRule.ConditionText)) {
            return;
        }

        if (rule is ICssFontFaceRule fontFace) {
            definitions.Add(new HtmlCssFontFaceDefinition(
                fontFace.Family,
                fontFace.Source,
                fontFace.Weight,
                fontFace.Style));
            return;
        }

        if (rule is ICssGroupingRule groupingRule) {
            foreach (ICssRule child in groupingRule.Rules) {
                AddFontFaces(child, mediaContext, definitions);
            }
        }
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

using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

public static partial class HtmlResourcePipeline {
    internal static HashSet<string> CollectProvenanceCssImageCustomProperties(IEnumerable<string> styles) {
        var usedImageProperties = new HashSet<string>(StringComparer.Ordinal);
        var dependencies = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
        foreach (string css in styles) {
            if (string.IsNullOrWhiteSpace(css)) continue;
            string masked = MaskCssComments(css);
            foreach (Match variable in CssVarExpression.Matches(masked)) {
                if (!IsCssFunctionNameAt(masked, variable.Index, "var") || IsInsideCssString(masked, variable.Index)) continue;
                string referencedProperty = DecodeCssEscapes(variable.Groups["name"].Value);
                if (ClassifyCssUrl(masked, variable.Index) == HtmlResourceKind.Image) usedImageProperties.Add(referencedProperty);
                if (!TryGetCustomPropertyName(masked, variable.Index, out string ownerProperty)) continue;
                ownerProperty = DecodeCssEscapes(ownerProperty);
                if (!dependencies.TryGetValue(ownerProperty, out HashSet<string>? referencedProperties)) {
                    referencedProperties = new HashSet<string>(StringComparer.Ordinal);
                    dependencies.Add(ownerProperty, referencedProperties);
                }
                referencedProperties.Add(referencedProperty);
            }
        }

        var pending = new Queue<string>(usedImageProperties);
        while (pending.Count > 0) {
            string property = pending.Dequeue();
            if (!dependencies.TryGetValue(property, out HashSet<string>? referencedProperties)) continue;
            foreach (string referencedProperty in referencedProperties) {
                if (usedImageProperties.Add(referencedProperty)) pending.Enqueue(referencedProperty);
            }
        }
        return usedImageProperties;
    }

    internal static IEnumerable<HtmlCssImageReference> EnumerateProvenanceCssImageReferences(
        string css,
        ISet<string>? documentUsedImageProperties = null) {
        if (string.IsNullOrWhiteSpace(css)) yield break;
        string masked = MaskCssComments(css);
        var usedImageProperties = documentUsedImageProperties == null
            ? new HashSet<string>(StringComparer.Ordinal)
            : new HashSet<string>(documentUsedImageProperties, StringComparer.Ordinal);
        foreach (Match variable in CssVarExpression.Matches(masked)) {
            if (IsCssFunctionNameAt(masked, variable.Index, "var") &&
                !IsInsideCssString(masked, variable.Index) &&
                ClassifyCssUrl(masked, variable.Index) == HtmlResourceKind.Image) {
                usedImageProperties.Add(DecodeCssEscapes(variable.Groups["name"].Value));
            }
        }
        var emittedRanges = new HashSet<(int Start, int Length)>();
        foreach (Match match in CssUrlExpression.Matches(masked)) {
            bool isCustomProperty = TryGetCustomPropertyName(masked, match.Index, out string customPropertyName);
            if (!IsCssFunctionNameAt(masked, match.Index, "url") ||
                IsInsideCssString(masked, match.Index) ||
                IsImportAtRuleUrl(masked, match.Index) ||
                IsAtRulePreludeUrl(masked, match.Index) ||
                isCustomProperty && !usedImageProperties.Contains(DecodeCssEscapes(customPropertyName)) ||
                !isCustomProperty && ClassifyCssUrl(masked, match.Index) != HtmlResourceKind.Image) continue;
            Group sourceGroup = match.Groups["url"];
            int leading = 0;
            while (leading < sourceGroup.Length && char.IsWhiteSpace(sourceGroup.Value[leading])) leading++;
            int trailing = sourceGroup.Length;
            while (trailing > leading && char.IsWhiteSpace(sourceGroup.Value[trailing - 1])) trailing--;
            if (trailing == leading) continue;
            string source = DecodeCssEscapes(sourceGroup.Value.Substring(leading, trailing - leading));
            var range = (sourceGroup.Index + leading, trailing - leading);
            if (emittedRanges.Add(range)) yield return new HtmlCssImageReference(range.Item1, range.Item2, source);
        }

        foreach (CssStringUrlReference reference in ExtractImageSetStringUrls(masked)) {
            if (ClassifyCssUrl(masked, reference.Start) != HtmlResourceKind.Image) continue;
            int start = css.IndexOf(reference.Source, reference.Start, Math.Min(css.Length - reference.Start, reference.End - reference.Start), StringComparison.Ordinal);
            if (start < 0 || !emittedRanges.Add((start, reference.Source.Length))) continue;
            yield return new HtmlCssImageReference(start, reference.Source.Length, DecodeCssEscapes(reference.Source));
        }
    }

    private static string MaskCssComments(string css) {
        var result = new StringBuilder(css);
        char quote = '\0';
        for (int index = 0; index < css.Length; index++) {
            char current = css[index];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, index)) quote = '\0';
                continue;
            }
            if (current is '"' or '\'') { quote = current; continue; }
            if (current != '/' || index + 1 >= css.Length || css[index + 1] != '*') continue;
            result[index++] = ' ';
            result[index] = ' ';
            while (index + 1 < css.Length && !(css[index] == '*' && css[index + 1] == '/')) result[++index] = ' ';
            if (index + 1 < css.Length) { result[index] = ' '; result[++index] = ' '; }
        }
        return result.ToString();
    }
}

internal readonly struct HtmlCssImageReference {
    internal HtmlCssImageReference(int start, int length, string value) {
        Start = start;
        Length = length;
        Value = value;
    }

    internal int Start { get; }
    internal int Length { get; }
    internal string Value { get; }
}

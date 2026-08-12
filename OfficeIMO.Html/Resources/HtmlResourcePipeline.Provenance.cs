using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

public static partial class HtmlResourcePipeline {
    internal static IEnumerable<HtmlCssImageReference> EnumerateProvenanceCssImageReferences(string css) {
        if (string.IsNullOrWhiteSpace(css)) yield break;
        string masked = MaskCssComments(css);
        foreach (Match match in CssUrlExpression.Matches(masked)) {
            if (!IsCssFunctionNameAt(masked, match.Index, "url") ||
                IsInsideCssString(masked, match.Index) ||
                IsImportAtRuleUrl(masked, match.Index) ||
                IsAtRulePreludeUrl(masked, match.Index) ||
                IsCustomPropertyUrl(masked, match.Index) ||
                ClassifyCssUrl(masked, match.Index) != HtmlResourceKind.Image) continue;
            Group sourceGroup = match.Groups["url"];
            int leading = 0;
            while (leading < sourceGroup.Length && char.IsWhiteSpace(sourceGroup.Value[leading])) leading++;
            int trailing = sourceGroup.Length;
            while (trailing > leading && char.IsWhiteSpace(sourceGroup.Value[trailing - 1])) trailing--;
            if (trailing == leading) continue;
            string source = DecodeCssEscapes(sourceGroup.Value.Substring(leading, trailing - leading));
            yield return new HtmlCssImageReference(sourceGroup.Index + leading, trailing - leading, source);
        }

        foreach (CssStringUrlReference reference in ExtractImageSetStringUrls(masked)) {
            if (ClassifyCssUrl(masked, reference.Start) != HtmlResourceKind.Image) continue;
            int start = css.IndexOf(reference.Source, reference.Start, Math.Min(css.Length - reference.Start, reference.End - reference.Start), StringComparison.Ordinal);
            if (start < 0) continue;
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

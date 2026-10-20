using System.Text;
using OfficeIMO.Html.Css;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    /// <summary>
    /// Keeps owned length/math values intact while the retained parser assigns
    /// shorthand components. Markers are local to this one declaration.
    /// </summary>
    private static string PreserveProviderLengths(string layer, out IReadOnlyDictionary<string, string> originals,
        Func<HtmlCssMathExpression, bool>? shouldPreserve = null) {
        var replacements = new Dictionary<string, string>(StringComparer.Ordinal);
        var result = new StringBuilder(layer.Length);
        IReadOnlyList<HtmlCssToken> tokens = HtmlCssTokenizer.Tokenize(layer, UnboundedPropertyTokenization);
        int copied = 0, skipThrough = -1, marker = 900001;
        foreach (HtmlCssToken token in tokens) {
            if (token.Offset <= skipThrough) continue;
            int end = token.Offset + token.Length;
            string value;
            if (token.Kind == HtmlCssTokenKind.Function) {
                int open = end - 1;
                if (!TryFindFunctionEnd(layer, open, out int close)) continue;
                skipThrough = close;
                if (!IsKnownKeyword(token.Value ?? "", "calc", "min", "max", "clamp")) continue;
                end = close + 1;
                value = layer.Substring(token.Offset, end - token.Offset);
            } else if (token.Kind == HtmlCssTokenKind.Dimension) value = token.GetText(layer);
            else continue;
            bool valid;
            try {
                HtmlCssMathParseResult parsed = HtmlCssMathParser.ParseLengthPercentage(value);
                valid = parsed.IsParsed && (shouldPreserve == null || shouldPreserve(parsed.Expression!));
            } catch (HtmlCssMathLimitException) { valid = false; }
            if (!valid) continue;
            string replacement;
            do { replacement = (marker++).ToString(System.Globalization.CultureInfo.InvariantCulture) + "px"; }
            while (layer.IndexOf(replacement, StringComparison.OrdinalIgnoreCase) >= 0);
            replacements.Add(replacement, value);
            result.Append(layer, copied, token.Offset - copied).Append(replacement);
            copied = end;
        }
        originals = replacements;
        if (copied == 0) return layer;
        return result.Append(layer, copied, layer.Length - copied).ToString();
    }

    private static string RestoreProviderLengths(string value, IReadOnlyDictionary<string, string> originals) {
        foreach (KeyValuePair<string, string> original in originals) value = value.Replace(original.Key, original.Value);
        return value;
    }
}

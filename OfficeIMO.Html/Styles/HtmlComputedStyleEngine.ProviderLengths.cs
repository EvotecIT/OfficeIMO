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
        bool smallDeclaration = layer.Length <= 4096 && tokens.Count <= 64;
        System.Collections.BitArray? occupiedMarkers = null;
        int copied = 0, skipThrough = -1, markerIndex = 0;
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
            if (smallDeclaration) {
                do { replacement = (900001 + markerIndex++).ToString(System.Globalization.CultureInfo.InvariantCulture) + "px"; }
                while (layer.IndexOf(replacement, StringComparison.OrdinalIgnoreCase) >= 0);
            } else {
                occupiedMarkers ??= OccupiedProviderLengthMarkers(layer);
                while (markerIndex < occupiedMarkers.Length && occupiedMarkers[markerIndex]) markerIndex++;
                if (markerIndex == occupiedMarkers.Length)
                    throw new InvalidOperationException("Background provider length marker budget exceeded.");
                replacement = (900001 + markerIndex).ToString(System.Globalization.CultureInfo.InvariantCulture) + "px";
                occupiedMarkers[markerIndex++] = true;
            }
            replacements.Add(replacement, value);
            result.Append(layer, copied, token.Offset - copied).Append(replacement);
            copied = end;
        }
        originals = replacements;
        if (copied == 0) return layer;
        return result.Append(layer, copied, layer.Length - copied).ToString();
    }

    private static System.Collections.BitArray OccupiedProviderLengthMarkers(string layer) {
        // Mark every numeric suffix immediately before "px", including text in quoted
        // values. This preserves the old substring-collision rule in one bounded scan.
        var occupied = new System.Collections.BitArray(1_000_000);
        for (int index = 1; index < layer.Length - 1; index++) {
            if ((layer[index] != 'p' && layer[index] != 'P') || (layer[index + 1] != 'x' && layer[index + 1] != 'X'))
                continue;
            int value = 0, place = 1;
            for (int digits = 0; digits < 7 && index > digits; digits++) {
                char digit = layer[index - digits - 1];
                if (digit < '0' || digit > '9') break;
                value += (digit - '0') * place;
                if (value >= 900001 && value <= 1900000) occupied[value - 900001] = true;
                place *= 10;
            }
        }
        return occupied;
    }

    private static string RestoreProviderLengths(string value, IReadOnlyDictionary<string, string> originals) {
        foreach (KeyValuePair<string, string> original in originals) value = value.Replace(original.Key, original.Value);
        return value;
    }
}

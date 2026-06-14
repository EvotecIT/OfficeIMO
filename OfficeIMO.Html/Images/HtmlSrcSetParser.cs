namespace OfficeIMO.Html;

/// <summary>
/// Parses HTML <c>srcset</c> attributes into URL and descriptor candidates.
/// </summary>
public static class HtmlSrcSetParser {
    /// <summary>
    /// Parses a <c>srcset</c> value while preserving candidate descriptors.
    /// </summary>
    public static IReadOnlyList<HtmlSrcSetCandidate> Parse(string? srcSet) {
        var candidates = new List<HtmlSrcSetCandidate>();
        if (string.IsNullOrWhiteSpace(srcSet)) {
            return candidates;
        }

        string value = srcSet!;
        int index = 0;
        while (index < value.Length) {
            SkipWhitespaceAndCommas(value, ref index);
            if (index >= value.Length) {
                break;
            }

            int urlStart = index;
            while (index < value.Length && !char.IsWhiteSpace(value[index])) {
                index++;
            }

            string url = value.Substring(urlStart, index - urlStart);
            int trailingCommaCount = 0;
            while (url.Length > 0 && url[url.Length - 1] == ',') {
                trailingCommaCount++;
                url = url.Substring(0, url.Length - 1);
            }

            url = url.Trim();
            if (url.Length == 0) {
                continue;
            }

            if (trailingCommaCount > 0) {
                candidates.Add(new HtmlSrcSetCandidate(url, string.Empty));
                continue;
            }

            SkipWhitespace(value, ref index);

            int descriptorStart = index;
            while (index < value.Length && value[index] != ',') {
                index++;
            }

            string descriptor = value.Substring(descriptorStart, index - descriptorStart).Trim();
            if (index < value.Length && value[index] == ',') {
                index++;
            }

            candidates.Add(new HtmlSrcSetCandidate(url, descriptor));
        }

        return candidates;
    }

    private static void SkipWhitespaceAndCommas(string value, ref int index) {
        while (index < value.Length && (char.IsWhiteSpace(value[index]) || value[index] == ',')) {
            index++;
        }
    }

    private static void SkipWhitespace(string value, ref int index) {
        while (index < value.Length && char.IsWhiteSpace(value[index])) {
            index++;
        }
    }
}

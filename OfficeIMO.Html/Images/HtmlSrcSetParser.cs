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
            while (index < value.Length
                   && !char.IsWhiteSpace(value[index])
                   && !IsCandidateSeparator(value, urlStart, index)) {
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

    private static bool IsCandidateSeparator(string value, int urlStart, int index) {
        if (value[index] != ',') {
            return false;
        }

        if (StartsWith(value, urlStart, "data:", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        if (Contains(value, urlStart, index, '?')) {
            return false;
        }

        int next = index + 1;
        while (next < value.Length && char.IsWhiteSpace(value[next])) {
            next++;
        }

        if (next >= value.Length) {
            return false;
        }

        int tokenEnd = next;
        while (tokenEnd < value.Length && !char.IsWhiteSpace(value[tokenEnd]) && value[tokenEnd] != ',') {
            tokenEnd++;
        }

        return LooksLikeUrlCandidate(value, next, tokenEnd);
    }

    private static bool LooksLikeUrlCandidate(string value, int startIndex, int endIndex) {
        if (startIndex >= endIndex) {
            return false;
        }

        if (StartsWith(value, startIndex, "http://", StringComparison.OrdinalIgnoreCase)
            || StartsWith(value, startIndex, "https://", StringComparison.OrdinalIgnoreCase)
            || StartsWith(value, startIndex, "data:", StringComparison.OrdinalIgnoreCase)
            || value[startIndex] == '/'
            || value[startIndex] == '.') {
            return true;
        }

        for (int i = startIndex; i < endIndex; i++) {
            if (value[i] == '.') {
                return i > startIndex && i + 1 < endIndex;
            }
        }

        return false;
    }

    private static bool Contains(string value, int startIndex, int endIndex, char search) {
        for (int i = startIndex; i < endIndex; i++) {
            if (value[i] == search) {
                return true;
            }
        }

        return false;
    }

    private static bool StartsWith(string value, int startIndex, string prefix, StringComparison comparison) {
        if (startIndex < 0 || startIndex + prefix.Length > value.Length) {
            return false;
        }

        return string.Compare(value, startIndex, prefix, 0, prefix.Length, comparison) == 0;
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

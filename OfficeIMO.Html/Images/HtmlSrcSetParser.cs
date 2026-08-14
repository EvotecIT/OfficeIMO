namespace OfficeIMO.Html;

/// <summary>
/// Parses HTML <c>srcset</c> attributes into URL and descriptor candidates.
/// </summary>
public static class HtmlSrcSetParser {
    /// <summary>
    /// Parses a <c>srcset</c> value while preserving candidate descriptors.
    /// </summary>
    public static IReadOnlyList<HtmlSrcSetCandidate> Parse(string? srcSet) {
        return Parse(srcSet, null);
    }

    /// <summary>
    /// Parses a <c>srcset</c> value while preserving candidate descriptors, stopping after the requested number of candidates.
    /// </summary>
    public static IReadOnlyList<HtmlSrcSetCandidate> Parse(string? srcSet, int? maxCandidates) {
        var candidates = new List<HtmlSrcSetCandidate>();
        foreach (HtmlSrcSetCandidate candidate in Enumerate(srcSet, maxCandidates)) {
            candidates.Add(candidate);
        }

        return candidates;
    }

    /// <summary>
    /// Enumerates a <c>srcset</c> value while preserving candidate descriptors.
    /// </summary>
    public static IEnumerable<HtmlSrcSetCandidate> Enumerate(string? srcSet) {
        return Enumerate(srcSet, null);
    }

    /// <summary>
    /// Enumerates a <c>srcset</c> value while preserving candidate descriptors, stopping after the requested number of candidates.
    /// </summary>
    public static IEnumerable<HtmlSrcSetCandidate> Enumerate(string? srcSet, int? maxCandidates) {
        if (string.IsNullOrEmpty(srcSet) || IsNonPositiveCandidateLimit(maxCandidates)) {
            yield break;
        }

        string value = srcSet!;
        int index = 0;
        int emittedCandidates = 0;
        while (index < value.Length) {
            SkipWhitespaceAndCommas(value, ref index);
            if (index >= value.Length) {
                break;
            }

            int urlStart = index;
            while (index < value.Length && !IsHtmlWhitespace(value[index])) {
                index++;
            }

            string url = value.Substring(urlStart, index - urlStart);
            int trailingCommaCount = 0;
            while (url.Length > 0 && url[url.Length - 1] == ',') {
                trailingCommaCount++;
                url = url.Substring(0, url.Length - 1);
            }

            url = TrimHtmlWhitespace(url);
            if (url.Length == 0) {
                continue;
            }

            if (trailingCommaCount > 0) {
                yield return new HtmlSrcSetCandidate(url, string.Empty);
                emittedCandidates++;
                if (HasReachedCandidateLimit(emittedCandidates, maxCandidates)) {
                    break;
                }

                continue;
            }

            SkipWhitespace(value, ref index);

            int descriptorStart = index;
            int parenthesesDepth = 0;
            while (index < value.Length) {
                char current = value[index];
                if (current == ',' && parenthesesDepth == 0) break;
                if (current == '(') parenthesesDepth++;
                else if (current == ')' && parenthesesDepth > 0) parenthesesDepth--;
                index++;
            }

            string descriptor = TrimHtmlWhitespace(value.Substring(descriptorStart, index - descriptorStart));
            if (index < value.Length && value[index] == ',') {
                index++;
            }

            yield return new HtmlSrcSetCandidate(url, descriptor);
            emittedCandidates++;
            if (HasReachedCandidateLimit(emittedCandidates, maxCandidates)) {
                break;
            }
        }
    }

    private static bool HasReachedCandidateLimit(int count, int? maxCandidates) {
        return maxCandidates.HasValue && count >= maxCandidates.Value;
    }

    private static bool IsNonPositiveCandidateLimit(int? maxCandidates) {
        return maxCandidates.HasValue && maxCandidates.Value <= 0;
    }

    private static void SkipWhitespaceAndCommas(string value, ref int index) {
        while (index < value.Length && (IsHtmlWhitespace(value[index]) || value[index] == ',')) {
            index++;
        }
    }

    private static void SkipWhitespace(string value, ref int index) {
        while (index < value.Length && IsHtmlWhitespace(value[index])) {
            index++;
        }
    }

    private static bool IsHtmlWhitespace(char value) => value is '\t' or '\n' or '\f' or '\r' or ' ';

    private static string TrimHtmlWhitespace(string value) {
        int start = 0;
        while (start < value.Length && IsHtmlWhitespace(value[start])) start++;
        int end = value.Length;
        while (end > start && IsHtmlWhitespace(value[end - 1])) end--;
        return start == 0 && end == value.Length ? value : value.Substring(start, end - start);
    }
}

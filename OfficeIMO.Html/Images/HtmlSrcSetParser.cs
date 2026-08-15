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
        int parsedCandidates = 0;
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
            parsedCandidates++;
            bool reachedCandidateLimit = HasReachedCandidateLimit(parsedCandidates, maxCandidates);

            if (trailingCommaCount > 0) {
                yield return new HtmlSrcSetCandidate(url, string.Empty, urlStart);
                if (reachedCandidateLimit) {
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

            if (IsValidDescriptorList(descriptor)) {
                yield return new HtmlSrcSetCandidate(url, descriptor, urlStart);
            }
            if (reachedCandidateLimit) break;
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

    private static bool IsValidDescriptorList(string descriptorList) {
        bool hasWidth = false;
        bool hasDensity = false;
        bool hasFutureHeight = false;
        foreach (string descriptor in SplitDescriptors(descriptorList)) {
            if (descriptor.EndsWith("w", StringComparison.Ordinal) &&
                TryParsePositiveInteger(descriptor.Substring(0, descriptor.Length - 1))) {
                if (hasWidth || hasDensity) return false;
                hasWidth = true;
            } else if (descriptor.EndsWith("x", StringComparison.Ordinal) &&
                TryParseNonNegativeFloatingPoint(descriptor.Substring(0, descriptor.Length - 1))) {
                if (hasWidth || hasDensity || hasFutureHeight) return false;
                hasDensity = true;
            } else if (descriptor.EndsWith("h", StringComparison.Ordinal) &&
                TryParsePositiveInteger(descriptor.Substring(0, descriptor.Length - 1))) {
                if (hasFutureHeight || hasDensity) return false;
                hasFutureHeight = true;
            } else {
                return false;
            }
        }
        return !hasFutureHeight || hasWidth;
    }

    private static IEnumerable<string> SplitDescriptors(string value) {
        int index = 0;
        while (index < value.Length) {
            while (index < value.Length && IsHtmlWhitespace(value[index])) index++;
            if (index >= value.Length) yield break;
            int start = index;
            int parenthesesDepth = 0;
            while (index < value.Length) {
                char current = value[index];
                if (current == '(') parenthesesDepth++;
                else if (current == ')' && parenthesesDepth > 0) parenthesesDepth--;
                else if (IsHtmlWhitespace(current) && parenthesesDepth == 0) break;
                index++;
            }
            yield return value.Substring(start, index - start);
        }
    }

    private static bool TryParsePositiveInteger(string value) {
        if (value.Length == 0) return false;
        ulong result = 0;
        foreach (char character in value) {
            if (character < '0' || character > '9') return false;
            ulong digit = (ulong)(character - '0');
            if (result > (ulong.MaxValue - digit) / 10UL) return false;
            result = result * 10UL + digit;
        }
        return result > 0;
    }

    private static bool TryParseNonNegativeFloatingPoint(string value) {
        if (!System.Text.RegularExpressions.Regex.IsMatch(
                value,
                "^-?(?:[0-9]+(?:\\.[0-9]+)?|\\.[0-9]+)(?:[eE][+-]?[0-9]+)?$",
                System.Text.RegularExpressions.RegexOptions.CultureInvariant,
                TimeSpan.FromMilliseconds(100)) ||
            !double.TryParse(
                value,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture,
                out double result)) return false;
        return result > 0D && !double.IsInfinity(result) && !double.IsNaN(result);
    }

    private static string TrimHtmlWhitespace(string value) {
        int start = 0;
        while (start < value.Length && IsHtmlWhitespace(value[start])) start++;
        int end = value.Length;
        while (end > start && IsHtmlWhitespace(value[end - 1])) end--;
        return start == 0 && end == value.Length ? value : value.Substring(start, end - start);
    }
}

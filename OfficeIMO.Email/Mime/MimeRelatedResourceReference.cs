namespace OfficeIMO.Email;

/// <summary>Matches exact related-resource references in HTML content.</summary>
internal static class MimeRelatedResourceReference {
    internal static bool ContainsContentId(string html, string contentId) {
        string expected = contentId.Trim().Trim('<', '>');
        if (expected.Length == 0) return false;

        int searchStart = 0;
        while (searchStart < html.Length) {
            int marker = html.IndexOf("cid:", searchStart, StringComparison.OrdinalIgnoreCase);
            if (marker < 0) return false;
            if (marker > 0 && IsSchemeCharacter(html[marker - 1])) {
                searchStart = marker + 4;
                continue;
            }

            int valueStart = marker + 4;
            int valueEnd = valueStart;
            while (valueEnd < html.Length && !IsCidTerminator(html[valueEnd])) valueEnd++;
            if (valueEnd > valueStart && MatchesDecoded(
                    html.Substring(valueStart, valueEnd - valueStart), expected,
                    StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
            searchStart = Math.Max(valueEnd, marker + 4);
        }
        return false;
    }

    internal static bool ContainsContentLocation(string html, string contentLocation) {
        string expected = DecodeReference(contentLocation.Trim());
        if (expected.Length == 0) return false;

        int searchStart = 0;
        while (searchStart < html.Length) {
            int equals = html.IndexOf('=', searchStart);
            if (equals < 0) break;
            if (!IsReferenceAttribute(html, equals)) {
                searchStart = equals + 1;
                continue;
            }
            int valueStart = SkipWhitespace(html, equals + 1);
            if (TryReadValue(html, valueStart, out string value, out int valueEnd) &&
                MatchesDecoded(value, expected, StringComparison.Ordinal)) {
                return true;
            }
            searchStart = Math.Max(valueEnd, equals + 1);
        }

        searchStart = 0;
        while (searchStart < html.Length) {
            int marker = html.IndexOf("url(", searchStart, StringComparison.OrdinalIgnoreCase);
            if (marker < 0) return false;
            int valueStart = SkipWhitespace(html, marker + 4);
            if (TryReadCssUrl(html, valueStart, out string value, out int valueEnd) &&
                MatchesDecoded(value, expected, StringComparison.Ordinal)) {
                return true;
            }
            searchStart = Math.Max(valueEnd, marker + 4);
        }
        return false;
    }

    private static bool TryReadValue(string html, int start, out string value, out int end) {
        value = string.Empty;
        end = start;
        if (start >= html.Length) return false;
        char quote = html[start] == '"' || html[start] == '\'' ? html[start++] : '\0';
        end = start;
        if (quote != '\0') {
            while (end < html.Length && html[end] != quote) end++;
        } else {
            while (end < html.Length && !char.IsWhiteSpace(html[end]) && html[end] != '>') end++;
        }
        if (end <= start) return false;
        value = html.Substring(start, end - start);
        if (end < html.Length) end++;
        return true;
    }

    private static bool TryReadCssUrl(string html, int start, out string value, out int end) {
        value = string.Empty;
        end = start;
        if (start >= html.Length) return false;
        char quote = html[start] == '"' || html[start] == '\'' ? html[start++] : '\0';
        end = start;
        if (quote != '\0') {
            while (end < html.Length && html[end] != quote) end++;
        } else {
            while (end < html.Length && html[end] != ')') end++;
        }
        int valueEnd = end;
        while (valueEnd > start && char.IsWhiteSpace(html[valueEnd - 1])) valueEnd--;
        if (valueEnd <= start) return false;
        value = html.Substring(start, valueEnd - start);
        while (end < html.Length && html[end] != ')') end++;
        if (end < html.Length) end++;
        return true;
    }

    private static bool MatchesDecoded(string value, string expected, StringComparison comparison) {
        return string.Equals(DecodeReference(value), expected, comparison);
    }

    private static string DecodeReference(string value) {
        string decoded = System.Net.WebUtility.HtmlDecode(value);
        try {
            return Uri.UnescapeDataString(decoded);
        } catch (UriFormatException) {
            return decoded;
        }
    }

    private static int SkipWhitespace(string value, int start) {
        while (start < value.Length && char.IsWhiteSpace(value[start])) start++;
        return start;
    }

    private static bool IsReferenceAttribute(string html, int equals) {
        int end = equals - 1;
        while (end >= 0 && char.IsWhiteSpace(html[end])) end--;
        int start = end;
        while (start >= 0 && (char.IsLetterOrDigit(html[start]) || html[start] == '-' || html[start] == ':')) {
            start--;
        }
        string name = html.Substring(start + 1, end - start);
        return string.Equals(name, "src", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "href", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "background", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "poster", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "data", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "action", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "formaction", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "cite", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "longdesc", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "usemap", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "srcset", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsSchemeCharacter(char value) {
        return char.IsLetterOrDigit(value) || value == '+' || value == '-' || value == '.';
    }

    private static bool IsCidTerminator(char value) {
        return char.IsWhiteSpace(value) || value == '"' || value == '\'' || value == '<' || value == '>' ||
            value == '(' || value == ')' || value == '[' || value == ']' || value == '{' || value == '}' ||
            value == '/';
    }
}

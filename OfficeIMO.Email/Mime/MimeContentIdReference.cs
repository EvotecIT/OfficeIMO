namespace OfficeIMO.Email;

/// <summary>Matches complete, percent-decoded cid URLs in HTML content.</summary>
internal static class MimeContentIdReference {
    internal static bool Contains(string html, string contentId) {
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
            while (valueEnd < html.Length && !IsUrlTerminator(html[valueEnd])) valueEnd++;
            if (valueEnd > valueStart) {
                string candidate = html.Substring(valueStart, valueEnd - valueStart);
                try {
                    candidate = Uri.UnescapeDataString(candidate);
                } catch (UriFormatException) {
                    // A malformed cid URL cannot reference a valid Content-ID.
                }
                if (string.Equals(candidate, expected, StringComparison.OrdinalIgnoreCase)) return true;
            }
            searchStart = Math.Max(valueEnd, marker + 4);
        }
        return false;
    }

    private static bool IsSchemeCharacter(char value) {
        return char.IsLetterOrDigit(value) || value == '+' || value == '-' || value == '.';
    }

    private static bool IsUrlTerminator(char value) {
        return char.IsWhiteSpace(value) || value == '"' || value == '\'' || value == '<' || value == '>' ||
            value == '(' || value == ')' || value == '[' || value == ']' || value == '{' || value == '}' ||
            value == '/';
    }
}

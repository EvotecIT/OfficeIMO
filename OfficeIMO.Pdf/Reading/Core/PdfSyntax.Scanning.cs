namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    private static int FindObjectEnd(string text, int start) {
        int searchFrom = start;
        while (searchFrom >= 0 && searchFrom < text.Length) {
            int streamIdx = IndexOfKeywordOutsideLiteralString(text, "stream", searchFrom, text.Length);
            int endObjIdx = IndexOfKeywordOutsideLiteralString(text, "endobj", searchFrom, text.Length);

            if (endObjIdx < 0) {
                return -1;
            }

            if (streamIdx < 0 || endObjIdx < streamIdx) {
                return endObjIdx + 6;
            }

            int afterStream = SkipEOL(text, streamIdx + 6, text.Length);
            int endStreamIdx = IndexOfKeyword(text, "endstream", afterStream, text.Length);
            if (endStreamIdx < 0) {
                return -1;
            }

            searchFrom = endStreamIdx + 9;
        }

        return -1;
    }

    private static int IndexOfKeywordOutsideLiteralString(string text, string keyword, int start, int limit) {
        if (start < 0) start = 0;
        if (limit > text.Length) limit = text.Length;

        int literalDepth = 0;
        bool escaped = false;
        for (int i = start; i < limit; i++) {
            char c = text[i];
            if (literalDepth > 0) {
                if (escaped) {
                    escaped = false;
                    continue;
                }

                if (c == '\\') {
                    escaped = true;
                    continue;
                }

                if (c == '(') {
                    literalDepth++;
                    continue;
                }

                if (c == ')') {
                    literalDepth--;
                }

                continue;
            }

            if (c == '(') {
                literalDepth = 1;
                continue;
            }

            if (i + keyword.Length <= limit &&
                string.CompareOrdinal(text, i, keyword, 0, keyword.Length) == 0 &&
                HasKeywordBoundary(text, i - 1, start, limit) &&
                HasKeywordBoundary(text, i + keyword.Length, start, limit)) {
                return i;
            }
        }

        return -1;
    }

    private static int FindDictEnd(string text, int dictStart, int limit) {
        int depth = 0;
        for (int i = dictStart; i + 1 < limit; i++) {
            char c = text[i]; char n = text[i + 1];
            if (c == '<' && n == '<') { depth++; i++; continue; }
            if (c == '>' && n == '>') { depth--; i++; if (depth == 0) return i + 1; continue; }
        }
        return -1;
    }

    private static int IndexOfKeyword(string text, string keyword, int start, int limit) {
        if (start < 0) start = 0;
        if (limit > text.Length) limit = text.Length;

        int searchFrom = start;
        while (searchFrom < limit) {
            int idx = text.IndexOf(keyword, searchFrom, StringComparison.Ordinal);
            if (idx < 0 || idx >= limit) {
                return -1;
            }

            int end = idx + keyword.Length;
            if (HasKeywordBoundary(text, idx - 1, start, limit) &&
                HasKeywordBoundary(text, end, start, limit)) {
                return idx;
            }

            searchFrom = idx + 1;
        }

        return -1;
    }

    private static bool HasKeywordBoundary(string text, int idx, int start, int limit) {
        if (idx < start || idx >= limit) {
            return true;
        }

        char c = text[idx];
        return char.IsWhiteSpace(c) || c is '<' or '>' or '[' or ']' or '%';
    }

    private static int SkipEOL(string text, int idx, int limit) {
        if (idx < limit) {
            if (text[idx] == '\r') idx++;
            if (idx < limit && text[idx] == '\n') idx++;
        }
        return idx;
    }

    private static string SafeSlice(string s, int start, int length, int maxLen) {
        int len = Math.Min(length, Math.Max(0, maxLen));
        if (start < 0) start = 0;
        if (start + len > s.Length) len = s.Length - start;
        if (len <= 0) return string.Empty;
        return s.Substring(start, len);
    }
}

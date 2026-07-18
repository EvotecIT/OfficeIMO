namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    private static int FindObjectEnd(string text, int start) {
        int searchFrom = start;
        while (searchFrom >= 0 && searchFrom < text.Length) {
            int streamIdx = IndexOfKeywordOutsideLiteralString(text, "stream", searchFrom, text.Length);
            int endObjIdx = IndexOfKeywordOutsideLiteralString(text, "endobj", searchFrom, text.Length);

            if (streamIdx < 0) {
                return endObjIdx < 0 ? -1 : endObjIdx + 6;
            }

            if (endObjIdx >= 0 && endObjIdx < streamIdx) {
                return endObjIdx + 6;
            }

            int afterStream = SkipEOL(text, streamIdx + 6, text.Length);
            int endStreamSearchFrom = afterStream;
            while (endStreamSearchFrom < text.Length) {
                int endStreamIdx = IndexOfKeyword(text, "endstream", endStreamSearchFrom, text.Length);
                if (endStreamIdx < 0) {
                    return -1;
                }

                int nextToken = SkipWhitespaceAndComments(text, endStreamIdx + 9, text.Length);
                if (IsKeywordAt(text, "endobj", nextToken, text.Length)) {
                    return nextToken + 6;
                }

                // Once a complete indirect-object header begins, this stream's endobj
                // boundary was genuinely absent. Do not consume a later object's endobj.
                if (IsIndirectObjectHeaderAt(text, nextToken)) {
                    return -1;
                }

                // Binary stream data may contain the bytes "endstream". Only the
                // delimiter immediately followed by endobj is structural.
                endStreamSearchFrom = endStreamIdx + 9;
            }
        }

        return -1;
    }

    private static int SkipWhitespaceAndComments(string text, int index, int limit) {
        while (index < limit) {
            while (index < limit && char.IsWhiteSpace(text[index])) {
                index++;
            }

            if (index >= limit || text[index] != '%') {
                return index;
            }

            while (index < limit && text[index] != '\r' && text[index] != '\n') {
                index++;
            }
        }

        return index;
    }

    private static bool IsKeywordAt(string text, string keyword, int index, int limit) =>
        index >= 0 &&
        index + keyword.Length <= limit &&
        string.CompareOrdinal(text, index, keyword, 0, keyword.Length) == 0 &&
        HasKeywordBoundary(text, index - 1, 0, limit) &&
        HasKeywordBoundary(text, index + keyword.Length, 0, limit);

    private static bool IsIndirectObjectHeaderAt(string text, int index) {
        var match = ObjRegex.Match(text, index);
        return match.Success && match.Index == index;
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

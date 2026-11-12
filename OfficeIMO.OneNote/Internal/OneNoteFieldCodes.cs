namespace OfficeIMO.OneNote;

/// <summary>
/// Resolves native OneNote field codes for display. OneNote stores a hyperlink as a protected run
/// <c>U+FDDF HYPERLINK "url"</c> followed by protected result runs; only the result is visible.
/// Projection removes the marker and instruction and carries a HYPERLINK target onto the result runs.
/// </summary>
internal static class OneNoteFieldCodes {
    internal const char FieldMarker = '﷟';
    private const string HyperlinkKeyword = "HYPERLINK";

    /// <summary>Returns display runs; the source runs are returned unchanged when they contain no field code.</summary>
    internal static IReadOnlyList<OneNoteTextRun> ResolveRuns(IList<OneNoteTextRun> runs) {
        if (!ContainsFieldCode(runs)) return runs as IReadOnlyList<OneNoteTextRun> ?? runs.ToArray();

        var result = new List<OneNoteTextRun>(runs.Count);
        string? resultLink = null;
        foreach (OneNoteTextRun run in runs) {
            string text = run.Text ?? string.Empty;
            if (run.MathExpression != null || text.IndexOf(FieldMarker) < 0) {
                if (resultLink != null && run.HyperlinkProtected && run.MathExpression == null) {
                    result.Add(Copy(run, text, run.Hyperlink ?? resultLink));
                } else {
                    resultLink = null;
                    result.Add(run);
                }
                continue;
            }

            int position = 0;
            string? link = resultLink;
            while (position <= text.Length) {
                int marker = text.IndexOf(FieldMarker, position);
                string segment = marker < 0 ? text.Substring(position) : text.Substring(position, marker - position);
                if (segment.Length > 0) result.Add(Copy(run, segment, run.Hyperlink ?? link));
                if (marker < 0) break;
                position = SkipInstruction(text, marker + 1, out link);
            }
            // A protected instruction run is followed by protected result runs that carry its target.
            resultLink = run.HyperlinkProtected ? link : null;
        }
        return result;
    }

    /// <summary>Removes field markers and HYPERLINK instructions from text that has lost its run boundaries.</summary>
    internal static string StripInstructions(string value) {
        if (value.IndexOf(FieldMarker) < 0) return value;
        var builder = new System.Text.StringBuilder(value.Length);
        int position = 0;
        while (position < value.Length) {
            int marker = value.IndexOf(FieldMarker, position);
            if (marker < 0) {
                builder.Append(value, position, value.Length - position);
                break;
            }
            builder.Append(value, position, marker - position);
            position = SkipInstruction(value, marker + 1, out string? link, consumeUnknownToEnd: false);
            _ = link;
        }
        return builder.ToString();
    }

    private static bool ContainsFieldCode(IList<OneNoteTextRun> runs) {
        for (int index = 0; index < runs.Count; index++) {
            if (runs[index].Text != null && runs[index].Text.IndexOf(FieldMarker) >= 0) return true;
        }
        return false;
    }

    // Returns the index after the instruction. A HYPERLINK instruction ends after its target and switches;
    // an unknown instruction occupies the rest of its run (native instructions sit in their own run).
    private static int SkipInstruction(string text, int start, out string? link, bool consumeUnknownToEnd = true) {
        link = null;
        int index = SkipSpaces(text, start);
        int keywordEnd = index;
        while (keywordEnd < text.Length && char.IsLetter(text[keywordEnd])) keywordEnd++;
        string keyword = text.Substring(index, keywordEnd - index);
        if (!string.Equals(keyword, HyperlinkKeyword, StringComparison.OrdinalIgnoreCase)) {
            return consumeUnknownToEnd ? text.Length : keywordEnd;
        }

        index = keywordEnd;
        while (true) {
            int tokenStart = SkipSpaces(text, index);
            if (tokenStart >= text.Length) return text.Length;
            if (text[tokenStart] == '\\') {
                // \l bookmark, \o tooltip and \t target take an argument; \m and \n do not.
                char name = tokenStart + 1 < text.Length ? char.ToLowerInvariant(text[tokenStart + 1]) : '\0';
                index = Math.Min(text.Length, tokenStart + 2);
                if (name == 'l' || name == 'o' || name == 't') index = ReadToken(text, SkipSpaces(text, index), out _);
                continue;
            }
            if (link != null) return index;
            int end = ReadToken(text, tokenStart, out string token);
            if (token.Length == 0) return end;
            link = token;
            index = end;
        }
    }

    private static int ReadToken(string text, int start, out string token) {
        if (start < text.Length && text[start] == '"') {
            int close = text.IndexOf('"', start + 1);
            if (close < 0) close = text.Length;
            token = text.Substring(start + 1, close - start - 1);
            return Math.Min(text.Length, close + 1);
        }
        int end = start;
        while (end < text.Length && !char.IsWhiteSpace(text[end]) && text[end] != FieldMarker) end++;
        token = text.Substring(start, end - start);
        return end;
    }

    private static int SkipSpaces(string text, int index) {
        while (index < text.Length && text[index] == ' ') index++;
        return index;
    }

    private static OneNoteTextRun Copy(OneNoteTextRun source, string text, string? hyperlink) {
        if (string.Equals(source.Text, text, StringComparison.Ordinal) && string.Equals(source.Hyperlink, hyperlink, StringComparison.Ordinal)) return source;
        OneNoteTextRun copy = OneNoteMathRunPreservation.Clone(source);
        copy.Text = text;
        copy.Hyperlink = string.IsNullOrWhiteSpace(hyperlink) ? null : hyperlink;
        return copy;
    }
}

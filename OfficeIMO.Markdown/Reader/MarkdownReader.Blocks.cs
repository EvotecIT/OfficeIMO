using System.Globalization;

namespace OfficeIMO.Markdown;

// Block parsing helpers for MarkdownReader.
internal static partial class MarkdownReader {
    private static bool IsAtxHeading(string line, out int level, out string text) {
        return TryGetAtxHeadingContentRange(line, out level, out _, out _, out text, out _, out _);
    }

    internal static bool TryGetSetextHeadingUnderlineLevel(string line, out int level) {
        level = 0;
        if (string.IsNullOrWhiteSpace(line)) return false;
        if (CountLeadingIndentColumns(line) > 3) return false;

        var trimmed = line.Trim();
        char marker = '\0';
        for (int i = 0; i < trimmed.Length; i++) {
            char ch = trimmed[i];
            if (ch != '=' && ch != '-') return false;
            if (marker == '\0') marker = ch;
            else if (ch != marker) return false;
        }

        level = marker == '=' ? 1 : 2;
        return true;
    }

    private static bool IsSetextHeadingUnderlineSuppressed(MarkdownReaderState? state, int zeroBasedLineIndex) =>
        state?.SuppressedSetextHeadingUnderlineLines.Contains(zeroBasedLineIndex) == true ||
        IsLazyQuoteContainerSetextUnderline(state, zeroBasedLineIndex);

    private static bool IsLazyQuoteContainerSetextUnderline(MarkdownReaderState? state, int zeroBasedLineIndex) {
        if (state == null ||
            !state.LazyQuoteContinuationLines.Contains(zeroBasedLineIndex) ||
            !state.QuoteContainerLines.Contains(zeroBasedLineIndex)) {
            return false;
        }

        for (var index = zeroBasedLineIndex - 1; index >= 0; index--) {
            if (state.QuoteContainerLines.Contains(index)) {
                return true;
            }

            if (!state.LazyQuoteContinuationLines.Contains(index)) {
                return false;
            }
        }

        return false;
    }

    private static bool TryGetAtxHeadingContentRange(
        string line,
        out int level,
        out int contentStart,
        out int contentEnd,
        out string text,
        out int closingMarkerStart,
        out int closingMarkerEnd) =>
        TryGetAtxHeadingContentRange(line, 0, line?.Length ?? 0, out level, out contentStart, out contentEnd, out text, out closingMarkerStart, out closingMarkerEnd);

    private static bool TryGetAtxHeadingContentRange(
        string line,
        int lineStart,
        int lineEnd,
        out int level,
        out int contentStart,
        out int contentEnd,
        out string text,
        out int closingMarkerStart,
        out int closingMarkerEnd) {
        level = 0;
        contentStart = 0;
        contentEnd = 0;
        closingMarkerStart = -1;
        closingMarkerEnd = -1;
        text = string.Empty;
        if (line == null || lineEnd <= lineStart) return false;
        int length = lineEnd - lineStart;

        int indent = 0;
        while (indent < length && indent < 4 && line[lineStart + indent] == ' ') indent++;
        if (indent >= 4) return false;

        int i = indent;
        while (i < length && line[lineStart + i] == '#') i++;

        int count = i - indent;
        if (count < 1 || count > 6) return false;
        if (i < length && !char.IsWhiteSpace(line[lineStart + i])) return false;

        contentStart = i;
        while (contentStart < length && char.IsWhiteSpace(line[lineStart + contentStart])) contentStart++;
        if (contentStart >= length) {
            level = count;
            text = string.Empty;
            contentEnd = contentStart;
            return true;
        }

        contentEnd = length;
        while (contentEnd > contentStart && char.IsWhiteSpace(line[lineStart + contentEnd - 1])) contentEnd--;

        int closingStart = contentEnd;
        int closingEnd = contentEnd;
        while (closingStart > contentStart && line[lineStart + closingStart - 1] == '#') closingStart--;
        if (closingStart < contentEnd) {
            int beforeClosing = closingStart - 1;
            if (beforeClosing < contentStart || char.IsWhiteSpace(line[lineStart + beforeClosing])) {
                closingMarkerStart = closingStart;
                closingMarkerEnd = closingEnd;
                contentEnd = beforeClosing < contentStart ? contentStart : beforeClosing;
                while (contentEnd > contentStart && char.IsWhiteSpace(line[lineStart + contentEnd - 1])) contentEnd--;
            }
        }

        level = count;
        text = line.Substring(lineStart + contentStart, contentEnd - contentStart);
        return true;
    }

}

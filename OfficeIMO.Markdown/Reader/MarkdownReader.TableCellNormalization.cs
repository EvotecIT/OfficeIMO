namespace OfficeIMO.Markdown;

internal static partial class MarkdownReader {
    private enum TableCellHtmlMode { Unrestricted, Parsed, Builder }

    /// <summary>Parses builder cell Markdown with the same literal handling as parsed tables.</summary>
    internal static InlineSequence ParseTableCellInlineText(string markdown, MarkdownLoadOptions? options, MarkdownReaderState? state) {
        var cell = new TableCellSourceFragment(markdown, markdown, default);
        var normalized = NormalizeTableCellInlineMarkdownWithSourceMap(cell, state: null, unescapeCodePipes: false).Text;
        return ParseInlines(normalized, options ?? new MarkdownLoadOptions(), state, tableCellHtmlMode: TableCellHtmlMode.Builder);
    }

    /// <summary>Normalizes cell break markers while preserving code spans and complete HTML tokens.</summary>
    internal static string NormalizeTableCellBreaks(string markdown) {
        var cell = new TableCellSourceFragment(markdown, markdown, default);
        return NormalizeTableCellInlineMarkdownWithSourceMap(cell, state: null, unescapeCodePipes: false, preserveFencedCodeSpans: false).Text;
    }

    private static bool TryConsumeTableCellLiteralToken(
        string text, int start, out int consumed, ref char blockFenceChar, ref int blockFenceLength,
        bool preserveFencedCodeSpans = true) {
        char marker = text[start];
        int runLength = 1;
        if (marker == '`' || marker == '~') {
            while (start + runLength < text.Length && text[start + runLength] == marker) runLength++;
            if (!preserveFencedCodeSpans && runLength >= 3) {
                if (blockFenceLength == 0) {
                    blockFenceChar = marker;
                    blockFenceLength = runLength;
                } else if (blockFenceChar == marker && runLength >= blockFenceLength) {
                    blockFenceLength = 0;
                }
                consumed = runLength;
                return true;
            }
        }

        // Break markers inside encoded fenced blocks represent code line endings, even between backticks.
        consumed = 0;
        if (blockFenceLength != 0) {
            return false;
        }
        if (marker == '`') {
            if (TryConsumeMatchedBacktickSpan(text, start, out consumed)) {
                return true;
            }
            consumed = runLength;
            return true;
        }
        return marker == '<'
            && !TableBlock.TryConsumeBreakTag(text, start, out _)
            && TryConsumeRawInlineHtmlTag(text, start, out consumed);
    }

    // Check tokens after the inline parser has separated code, link metadata and raw HTML.

    private static bool IsTableCellInlineHtmlAllowed(string html, TableCellHtmlMode mode) {
        if (mode == TableCellHtmlMode.Unrestricted) {
            return true;
        }

        if (TableBlock.TryConsumeSupportedInlineFormattingTag(html, 0, out int formattingLength)
            && formattingLength == html.Length) {
            return true;
        }

        return mode == TableCellHtmlMode.Parsed
            && TryConsumeInlineRawHtmlTableCellTag(html, 0, out int rawLength)
            && rawLength == html.Length;
    }

    private static (string Text, MarkdownInlineSourceMap? SourceMap) NormalizeTableCellInlineMarkdownWithSourceMap(
        TableCellSourceFragment cell,
        MarkdownReaderState? state,
        bool unescapeCodePipes = true,
        bool preserveFencedCodeSpans = true) {
        if (string.IsNullOrEmpty(cell.Markdown)) {
            return (string.Empty, null);
        }

        var builder = new StringBuilder(cell.Markdown.Length + 8);
        List<MarkdownSourcePoint?>? points = state?.SourceTextMap == null ? null : new List<MarkdownSourcePoint?>(cell.Markdown.Length + 8);
        List<MarkdownSourceSpan?>? tokenSpans = points == null ? null : new List<MarkdownSourceSpan?>(cell.Markdown.Length + 8);
        List<string?>? tokenLiterals = points == null ? null : new List<string?>(cell.Markdown.Length + 8);
        int absoluteLine = cell.SourceSpan.StartLine;
        int column = cell.SourceSpan.StartColumn ?? 1;
        char blockFenceChar = '\0';
        int blockFenceLength = 0;

        void AppendMapped(string value, int sourceColumn, MarkdownSourceSpan? tokenSpan = null, string? tokenLiteral = null) {
            builder.Append(value);
            if (points == null) {
                return;
            }

            var point = state!.SourceTextMap!.CreatePoint(absoluteLine, sourceColumn);
            for (int i = 0; i < value.Length; i++) {
                points.Add(point);
                tokenSpans!.Add(i == 0 ? tokenSpan : null);
                tokenLiterals!.Add(i == 0 ? tokenLiteral : null);
            }
        }

        for (int i = 0; i < cell.Markdown.Length; i++) {
            char ch = cell.Markdown[i];
            if (TryConsumeTableCellLiteralToken(cell.Markdown, i, out int literalLength, ref blockFenceChar, ref blockFenceLength, preserveFencedCodeSpans)) {
                for (int offset = 0; offset < literalLength; offset++) {
                    char literal = cell.Markdown[i + offset];
                    if (unescapeCodePipes && ch == '`' && literal == '\\'
                        && offset + 1 < literalLength && cell.Markdown[i + offset + 1] == '|') {
                        offset++;
                        literal = '|';
                    }
                    AppendMapped(literal.ToString(), column + offset);
                }
                column += literalLength;
                i += literalLength - 1;
                continue;
            }

            if (ch == '\\' && i + 1 < cell.Markdown.Length && IsBackslashEscapable(cell.Markdown[i + 1])) {
                AppendMapped(ch.ToString(), column++);
                AppendMapped(cell.Markdown[++i].ToString(), column++);
                continue;
            }

            if (ch == '\r') {
                if (i + 1 < cell.Markdown.Length && cell.Markdown[i + 1] == '\n') {
                    i++;
                }
                AppendMapped("\n", column);
                column++;
                continue;
            }

            if (ch == '\n') {
                AppendMapped("\n", column);
                column++;
                continue;
            }

            if (ch == '<' && TableBlock.TryConsumeBreakTag(cell.Markdown, i, out int consumed)) {
                AppendMapped(
                    "\n",
                    column,
                    CreateSpan(state, absoluteLine, column, absoluteLine, column + consumed - 1),
                    cell.Markdown.Substring(i, consumed));
                column += consumed;
                i += consumed - 1;
                continue;
            }

            AppendMapped(ch.ToString(), column);
            column++;
        }

        return (builder.ToString(), points == null ? null : new MarkdownInlineSourceMap(points.ToArray(), tokenSpans!.ToArray(), tokenLiterals!.ToArray()));
    }

}

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static bool TryParseSupportedInlineHtmlTag(
        string text,
        int start,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        MarkdownInlineSourceMap? sourceMap,
        bool allowLinks,
        bool allowImages,
        out int consumed,
        out IMarkdownInline htmlNode) {
        consumed = 0;
        htmlNode = null!;

        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length || text[start] != '<') {
            return false;
        }

        string[] tags = { "u", "sup", "sub", "ins", "q" };
        for (int i = 0; i < tags.Length; i++) {
            if (!TryParseInlineHtmlWrapper(text, start, tags[i], options, state, sourceMap, allowLinks, allowImages, out consumed, out var inlines)) {
                continue;
            }

            var htmlTag = new HtmlTagSequenceInline(tags[i], inlines);
            SetInlineHtmlTagMarkerSpans(htmlTag, sourceMap, start, consumed);
            htmlNode = htmlTag;
            return true;
        }

        return false;
    }

    private static bool TryParseInlineHtmlWrapper(
        string text,
        int start,
        string tagName,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        MarkdownInlineSourceMap? sourceMap,
        bool allowLinks,
        bool allowImages,
        out int consumed,
        out InlineSequence inlines) {
        consumed = 0;
        inlines = new InlineSequence();

        if (!StartsWithExactHtmlTag(text, start, tagName, opening: true)) {
            return false;
        }

        int openLength = tagName.Length + 2;
        int scan = start + openLength;
        int depth = 1;

        while (scan < text.Length) {
            if (StartsWithExactHtmlTag(text, scan, tagName, opening: false)) {
                depth--;
                if (depth == 0) {
                    string inner = text.Substring(start + openLength, scan - (start + openLength));
                    inlines = ParseInlinesInternal(
                        inner,
                        options,
                        state,
                        allowLinks,
                        allowImages,
                        sourceMap?.Slice(start + openLength, inner.Length));
                    DecodeHtmlEntitiesInTextRuns(inlines);
                    consumed = (scan - start) + tagName.Length + 3;
                    return true;
                }

                scan += tagName.Length + 3;
                continue;
            }

            if (StartsWithExactHtmlTag(text, scan, tagName, opening: true)) {
                depth++;
                scan += openLength;
                continue;
            }

            scan++;
        }

        return false;
    }

    private static void SetInlineHtmlTagMarkerSpans(
        HtmlTagSequenceInline htmlTag,
        MarkdownInlineSourceMap? sourceMap,
        int start,
        int consumed) {
        if (htmlTag == null || sourceMap == null || consumed <= 0) {
            return;
        }

        var openingMarker = "<" + htmlTag.TagName + ">";
        var closingMarker = "</" + htmlTag.TagName + ">";
        MarkdownInlineMetadataSourceSpans.SetFormattingMarkers(
            htmlTag,
            openingMarker,
            sourceMap.GetSpan(start, openingMarker.Length),
            closingMarker,
            sourceMap.GetSpan(start + consumed - closingMarker.Length, closingMarker.Length));
    }

    private static bool DecodeHtmlEntitiesInTextRuns(InlineSequence sequence) {
        if (sequence == null || sequence.Nodes.Count == 0) {
            return false;
        }

        var rewritten = new List<IMarkdownInline>(sequence.Nodes.Count);
        bool changed = false;

        for (int i = 0; i < sequence.Nodes.Count; i++) {
            var node = sequence.Nodes[i];
            if (node == null) {
                continue;
            }

            rewritten.Add(DecodeHtmlEntitiesInInlineNode(node, ref changed));
        }

        if (changed) {
            sequence.ReplaceItems(rewritten);
        }

        return changed;
    }

    private static IMarkdownInline DecodeHtmlEntitiesInInlineNode(IMarkdownInline node, ref bool changed) {
        if (node is TextRun text) {
            string decoded = System.Net.WebUtility.HtmlDecode(text.Text);
            if (!string.Equals(decoded, text.Text, StringComparison.Ordinal)) {
                changed = true;
                var decodedRun = new DecodedHtmlEntityTextRun(decoded);
                var sourceSpan = MarkdownInlineSourceSpans.Get(text);
                MarkdownInlineSourceSpans.Set(decodedRun, sourceSpan);
                MarkdownInlineMetadataSourceSpans.SetDecodedEntity(decodedRun, text.Text, sourceSpan);
                return decodedRun;
            }

            return text;
        }

        if (node is IInlineContainerMarkdownInline container && container.NestedInlines != null) {
            if (DecodeHtmlEntitiesInTextRuns(container.NestedInlines)) {
                changed = true;
            }
        }

        return node;
    }

    private static bool TryConsumeHtmlEntityText(string text, int start, out int consumed, out string decoded) {
        consumed = 0;
        decoded = string.Empty;

        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length || text[start] != '&') {
            return false;
        }

        int semicolon = text.IndexOf(';', start + 1);
        if (semicolon < 0 || semicolon == start + 1 || semicolon - start > 32) {
            return false;
        }

        string candidate = text.Substring(start, semicolon - start + 1);
        string htmlDecoded = System.Net.WebUtility.HtmlDecode(candidate);
        if (string.Equals(htmlDecoded, candidate, StringComparison.Ordinal)) {
            return false;
        }

        consumed = candidate.Length;
        decoded = htmlDecoded;
        return true;
    }

    private static bool TryConsumeRawInlineHtmlTag(string text, int start, out int consumed) {
        return TryConsumeRawInlineHtmlTag(text, start, null, out consumed, out _);
    }

    private static bool TryConsumeRawInlineHtmlTag(
        string text,
        int start,
        MarkdownInlineSourceMap? sourceMap,
        out int consumed,
        out string html) {
        consumed = 0;
        html = string.Empty;

        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length || text[start] != '<') {
            return false;
        }

        int position = start + 1;
        if (position >= text.Length) {
            return false;
        }

        if (text[position] == '/') {
            position++;
            if (position >= text.Length || !IsAsciiLetter(text[position])) {
                return false;
            }

            position = ConsumeHtmlTagName(text, position);
            while (position < text.Length && IsHtmlAttributeWhitespace(text[position])) {
                position++;
            }

            if (position < text.Length && text[position] == '>') {
                consumed = position - start + 1;
                html = RestoreRawInlineHtmlLiteral(text, start, consumed, sourceMap);
                return true;
            }

            return false;
        }

        if (!IsAsciiLetter(text[position])) {
            return false;
        }

        position = ConsumeHtmlTagName(text, position);
        if (position >= text.Length) {
            return false;
        }

        char next = text[position];
        if (next != '>' && next != '/' && !IsHtmlAttributeWhitespace(next)) {
            return false;
        }

        char quote = '\0';
        while (position < text.Length) {
            char ch = text[position];

            if (quote != '\0') {
                if (ch == quote) {
                    quote = '\0';
                }

                position++;
                continue;
            }

            if (ch == '"' || ch == '\'') {
                quote = ch;
                position++;
                continue;
            }

            if (ch == '\r' || ch == '\n' || ch == '<') {
                return false;
            }

            if (ch == '>') {
                consumed = position - start + 1;
                html = RestoreRawInlineHtmlLiteral(text, start, consumed, sourceMap);
                return true;
            }

            position++;
        }

        return false;
    }

    private static string RestoreRawInlineHtmlLiteral(string text, int start, int consumed, MarkdownInlineSourceMap? sourceMap) =>
        sourceMap?.RestoreSourceLineBreaks(text, start, consumed) ?? text.Substring(start, consumed);

    private static int ConsumeHtmlTagName(string text, int position) {
        while (position < text.Length) {
            char ch = text[position];
            if (!IsAsciiLetter(ch) && !char.IsDigit(ch) && ch != '-') {
                break;
            }

            position++;
        }

        return position;
    }

    private static bool IsAsciiLetter(char value) =>
        (value >= 'A' && value <= 'Z') || (value >= 'a' && value <= 'z');

    private static bool IsHtmlAttributeWhitespace(char value) =>
        value == ' ' || value == '\t';

    private static bool StartsWithExactHtmlTag(string text, int start, string tagName, bool opening) {
        if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(tagName) || start < 0 || start >= text.Length || text[start] != '<') {
            return false;
        }

        int position = start + 1;
        if (!opening) {
            if (position >= text.Length || text[position] != '/') {
                return false;
            }
            position++;
        }

        if (position + tagName.Length >= text.Length) {
            return false;
        }

        if (string.Compare(text, position, tagName, 0, tagName.Length, StringComparison.OrdinalIgnoreCase) != 0) {
            return false;
        }

        position += tagName.Length;
        if (position >= text.Length || text[position] != '>') {
            return false;
        }

        return true;
    }
}

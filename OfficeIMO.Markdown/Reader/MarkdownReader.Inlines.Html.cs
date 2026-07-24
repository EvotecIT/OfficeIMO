namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static readonly string[] SupportedInlineHtmlWrapperTags = { "u", "sup", "sub", "ins", "q" };
    private const int MaximumInlineHtmlWrapperDepth = 32;

    private sealed class InlineHtmlWrapperMatchIndex {
        private readonly Dictionary<int, (string TagName, int ClosingStart)> _matches;
        private readonly int _baseOffset;

        internal InlineHtmlWrapperMatchIndex(Dictionary<int, (string TagName, int ClosingStart)> matches, int baseOffset = 0) {
            _matches = matches;
            _baseOffset = baseOffset;
        }

        internal InlineHtmlWrapperMatchIndex Slice(int relativeOffset) =>
            new(_matches, checked(_baseOffset + relativeOffset));

        internal bool TryGet(int relativeStart, out string tagName, out int closingStart) {
            if (_matches.TryGetValue(checked(_baseOffset + relativeStart), out var match)) {
                tagName = match.TagName;
                closingStart = match.ClosingStart - _baseOffset;
                return true;
            }

            tagName = string.Empty;
            closingStart = -1;
            return false;
        }
    }

    private static InlineHtmlWrapperMatchIndex BuildInlineHtmlWrapperMatchIndex(string text) {
        var matches = new Dictionary<int, (string TagName, int ClosingStart)>();
        if (string.IsNullOrEmpty(text)) {
            return new InlineHtmlWrapperMatchIndex(matches);
        }

        var openingsByTag = new Dictionary<string, Stack<int>>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < SupportedInlineHtmlWrapperTags.Length; i++) {
            openingsByTag[SupportedInlineHtmlWrapperTags[i]] = new Stack<int>();
        }

        for (int position = 0; position < text.Length; position++) {
            if (text[position] != '<') {
                continue;
            }

            for (int tagIndex = 0; tagIndex < SupportedInlineHtmlWrapperTags.Length; tagIndex++) {
                string tagName = SupportedInlineHtmlWrapperTags[tagIndex];
                if (StartsWithExactHtmlTag(text, position, tagName, opening: true)) {
                    openingsByTag[tagName].Push(position);
                    position += tagName.Length + 1;
                    break;
                }

                if (StartsWithExactHtmlTag(text, position, tagName, opening: false)) {
                    var openings = openingsByTag[tagName];
                    if (openings.Count > 0) {
                        matches[openings.Pop()] = (tagName, position);
                    }

                    position += tagName.Length + 2;
                    break;
                }
            }
        }

        return new InlineHtmlWrapperMatchIndex(matches);
    }

    private static bool TryParseSupportedInlineHtmlTag(
        string text,
        int start,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        MarkdownInlineSourceMap? sourceMap,
        bool allowLinks,
        bool allowImages,
        InlineHtmlWrapperMatchIndex matches,
        int wrapperDepth,
        out int consumed,
        out IMarkdownInline htmlNode) {
        consumed = 0;
        htmlNode = null!;

        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length || text[start] != '<') {
            return false;
        }

        if (wrapperDepth >= MaximumInlineHtmlWrapperDepth
            || !matches.TryGet(start, out string tagName, out int closingStart)) {
            return false;
        }

        int openLength = tagName.Length + 2;
        int closingLength = tagName.Length + 3;
        if (closingStart < start + openLength || closingStart > text.Length - closingLength) {
            return false;
        }

        string inner = text.Substring(start + openLength, closingStart - (start + openLength));
        var inlines = ParseInlinesInternal(
            inner,
            options,
            state,
            allowLinks,
            allowImages,
            sourceMap?.Slice(start + openLength, inner.Length),
            matches.Slice(start + openLength),
            wrapperDepth + 1);
        DecodeHtmlEntitiesInTextRuns(inlines);
        consumed = (closingStart - start) + tagName.Length + 3;
        var htmlTag = new HtmlTagSequenceInline(tagName, inlines);
        SetInlineHtmlTagMarkerSpans(htmlTag, sourceMap, start, consumed);
        htmlNode = htmlTag;
        return true;
    }

    private static bool TryConsumeSupportedInlineHtmlWrapperSpan(
        string text,
        int start,
        out int consumed,
        InlineHtmlWrapperMatchIndex? matches = null) {
        consumed = 0;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length || text[start] != '<') {
            return false;
        }

        matches ??= BuildInlineHtmlWrapperMatchIndex(text);
        if (!matches.TryGet(start, out string tagName, out int closingStart)) {
            return false;
        }

        int openLength = tagName.Length + 2;
        int closingLength = tagName.Length + 3;
        if (closingStart < start + openLength || closingStart > text.Length - closingLength) {
            return false;
        }

        consumed = (closingStart - start) + tagName.Length + 3;
        return true;
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
            string decoded = CommonMarkCharacterReference.DecodeAll(text.Text);
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
        return CommonMarkCharacterReference.TryDecode(text, start, out consumed, out decoded);
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

        if (TryConsumeRawInlineHtmlComment(text, start, sourceMap, out consumed, out html)) {
            return true;
        }

        if (TryConsumeRawInlineHtmlProcessingInstruction(text, start, sourceMap, out consumed, out html)) {
            return true;
        }

        if (TryConsumeRawInlineHtmlDeclaration(text, start, sourceMap, out consumed, out html)) {
            return true;
        }

        if (TryConsumeRawInlineHtmlCData(text, start, sourceMap, out consumed, out html)) {
            return true;
        }

        string candidate = text.Substring(start);
        if (!HtmlBlockParser.TryParseHtmlTag(candidate, out _, out _, out int endIndex)) {
            return false;
        }

        consumed = endIndex + 1;
        html = RestoreRawInlineHtmlLiteral(text, start, consumed, sourceMap);
        return true;
    }

    private static bool TryConsumeRawInlineHtmlComment(
        string text,
        int start,
        MarkdownInlineSourceMap? sourceMap,
        out int consumed,
        out string html) {
        consumed = 0;
        html = string.Empty;

        if (!StartsWithOrdinal(text, start, "<!--")) {
            return false;
        }

        if (StartsWithOrdinal(text, start, "<!-->")) {
            consumed = 5;
            html = RestoreRawInlineHtmlLiteral(text, start, consumed, sourceMap);
            return true;
        }

        if (StartsWithOrdinal(text, start, "<!--->")) {
            consumed = 6;
            html = RestoreRawInlineHtmlLiteral(text, start, consumed, sourceMap);
            return true;
        }

        int end = text.IndexOf("-->", start + 4, StringComparison.Ordinal);
        if (end < 0) {
            return false;
        }

        consumed = end - start + 3;
        html = RestoreRawInlineHtmlLiteral(text, start, consumed, sourceMap);
        return true;
    }

    private static bool TryConsumeRawInlineHtmlProcessingInstruction(
        string text,
        int start,
        MarkdownInlineSourceMap? sourceMap,
        out int consumed,
        out string html) {
        consumed = 0;
        html = string.Empty;

        if (!StartsWithOrdinal(text, start, "<?")) {
            return false;
        }

        int end = text.IndexOf("?>", start + 2, StringComparison.Ordinal);
        if (end < 0) {
            return false;
        }

        consumed = end - start + 2;
        html = RestoreRawInlineHtmlLiteral(text, start, consumed, sourceMap);
        return true;
    }

    private static bool TryConsumeRawInlineHtmlDeclaration(
        string text,
        int start,
        MarkdownInlineSourceMap? sourceMap,
        out int consumed,
        out string html) {
        consumed = 0;
        html = string.Empty;

        if (start + 2 >= text.Length || text[start] != '<' || text[start + 1] != '!' || !IsAsciiUppercaseLetter(text[start + 2])) {
            return false;
        }

        int end = text.IndexOf('>', start + 3);
        if (end < 0) {
            return false;
        }

        consumed = end - start + 1;
        html = RestoreRawInlineHtmlLiteral(text, start, consumed, sourceMap);
        return true;
    }

    private static bool TryConsumeRawInlineHtmlCData(
        string text,
        int start,
        MarkdownInlineSourceMap? sourceMap,
        out int consumed,
        out string html) {
        consumed = 0;
        html = string.Empty;

        if (!StartsWithOrdinal(text, start, "<![CDATA[")) {
            return false;
        }

        int end = text.IndexOf("]]>", start + 9, StringComparison.Ordinal);
        if (end < 0) {
            return false;
        }

        consumed = end - start + 3;
        html = RestoreRawInlineHtmlLiteral(text, start, consumed, sourceMap);
        return true;
    }

    private static bool StartsWithOrdinal(string text, int start, string value) =>
        !string.IsNullOrEmpty(text)
        && !string.IsNullOrEmpty(value)
        && start >= 0
        && start <= text.Length - value.Length
        && string.Compare(text, start, value, 0, value.Length, StringComparison.Ordinal) == 0;

    private static string RestoreRawInlineHtmlLiteral(string text, int start, int consumed, MarkdownInlineSourceMap? sourceMap) =>
        sourceMap?.RestoreSourceLineBreaks(text, start, consumed) ?? text.Substring(start, consumed);

    private static bool IsAsciiLetter(char value) =>
        (value >= 'A' && value <= 'Z') || (value >= 'a' && value <= 'z');

    private static bool IsAsciiUppercaseLetter(char value) =>
        value >= 'A' && value <= 'Z';

    private static bool IsHtmlAttributeWhitespace(char value) =>
        value == ' ' || value == '\t' || value == '\n' || value == '\r' || value == '\f';

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

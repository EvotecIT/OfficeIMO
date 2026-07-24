namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static int FindMatchingBracket(
        string text,
        int openIndex,
        MarkdownReaderOptions? options = null,
        InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches = null) {
        if (string.IsNullOrEmpty(text) || openIndex < 0 || openIndex >= text.Length || text[openIndex] != '[') return -1;

        int depth = 0;
        bool escaped = false;
        var effectiveInlineHtmlWrapperMatches = inlineHtmlWrapperMatches;
        for (int i = openIndex; i < text.Length; i++) {
            char c = text[i];
            if (escaped) {
                escaped = false;
                continue;
            }

            if (c == '\\') {
                escaped = true;
                continue;
            }

            if (c == '<' && effectiveInlineHtmlWrapperMatches == null && options?.InlineHtml != false) {
                effectiveInlineHtmlWrapperMatches = BuildInlineHtmlWrapperMatchIndex(text);
            }

            if (TrySkipLinkLabelInlineSpan(text, i, options, out int spanConsumed, effectiveInlineHtmlWrapperMatches)) {
                i += spanConsumed - 1;
                continue;
            }

            if (c == '[') {
                depth++;
                continue;
            }

            if (c == ']') {
                depth--;
                if (depth == 0) return i;
            }
        }

        return -1;
    }

    private static bool TrySkipLinkLabelInlineSpan(
        string text,
        int start,
        MarkdownReaderOptions? options,
        out int consumed,
        InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches = null) {
        consumed = 0;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return false;

        if (text[start] == '`' && TryConsumeMatchedBacktickSpan(text, start, out consumed)) {
            return true;
        }

        if (text[start] == '<') {
            bool inlineHtml = options?.InlineHtml != false;
            if (inlineHtml && TryConsumeSupportedInlineHtmlWrapperSpan(text, start, out consumed, inlineHtmlWrapperMatches)) {
                return true;
            }

            if (TryParseAngleAutolink(text, start, out consumed, out _, out _)) {
                return true;
            }

            if (inlineHtml && TryConsumeRawInlineHtmlTag(text, start, out consumed)) {
                return true;
            }
        }

        return false;
    }

    private static bool TryConsumeMatchedBacktickSpan(string text, int start, out int consumed) {
        consumed = 0;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length || text[start] != '`') {
            return false;
        }

        int fenceLength = 0;
        while (start + fenceLength < text.Length && text[start + fenceLength] == '`') {
            fenceLength++;
        }

        int scan = start + fenceLength;
        while (scan < text.Length) {
            if (text[scan] != '`') {
                scan++;
                continue;
            }

            int runLength = 0;
            while (scan + runLength < text.Length && text[scan + runLength] == '`') {
                runLength++;
            }

            if (runLength == fenceLength) {
                consumed = scan + runLength - start;
                return true;
            }

            scan += runLength;
        }

        return false;
    }

    private static int FindReferenceLabelEnd(string text, int openIndex) {
        if (string.IsNullOrEmpty(text) || openIndex < 0 || openIndex >= text.Length || text[openIndex] != '[') return -1;

        bool escaped = false;
        for (int i = openIndex + 1; i < text.Length; i++) {
            char c = text[i];
            if (escaped) {
                escaped = false;
                continue;
            }

            if (c == '\\') {
                escaped = true;
                continue;
            }

            if (c == '[') return -1;
            if (c == ']') return i;
        }

        return -1;
    }

    private static string UnescapeMarkdownBackslashEscapes(string value) {
        if (string.IsNullOrEmpty(value)) return value ?? string.Empty;

        var sb = new StringBuilder(value.Length);
        for (int i = 0; i < value.Length; i++) {
            char c = value[i];
            if (c == '\\' && i + 1 < value.Length && IsBackslashEscapable(value[i + 1])) {
                sb.Append(value[i + 1]);
                i++;
                continue;
            }

            sb.Append(c);
        }

        return sb.ToString();
    }

    private static bool ContainsBackslashEscapableCharacter(string value) {
        if (string.IsNullOrEmpty(value)) return false;

        for (int i = 0; i + 1 < value.Length; i++) {
            if (value[i] == '\\' && IsBackslashEscapable(value[i + 1])) {
                return true;
            }
        }

        return false;
    }

    private static string DecodeLinkDestinationOrTitle(string value) {
        var unescaped = UnescapeMarkdownBackslashEscapes(value);
        return CommonMarkCharacterReference.DecodeAll(unescaped);
    }

    private static string DecodeLinkDestination(string value) {
        var unescaped = UnescapeMarkdownDestination(value);
        return CommonMarkCharacterReference.DecodeAll(unescaped);
    }

    private static string UnescapeMarkdownDestination(string value) {
        if (IsWindowsDriveLike(value)) return value.Replace("\\\\", "\\");
        return UnescapeMarkdownBackslashEscapes(value);
    }

    private static bool TryParseRefLink(string text, int start, MarkdownReaderOptions? options, out int consumed, out string label, out string refLabel, InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches = null) {
        consumed = 0; label = refLabel = string.Empty;
        if (start >= text.Length || text[start] != '[') return false;
        int rb = FindMatchingBracket(text, start, options, inlineHtmlWrapperMatches); if (rb < 0) return false;
        if (rb + 1 >= text.Length || text[rb + 1] != '[') return false;
        int rb2 = FindMatchingBracket(text, rb + 1, options, inlineHtmlWrapperMatches); if (rb2 < 0) return false;
        label = text.Substring(start + 1, rb - (start + 1));
        refLabel = text.Substring(rb + 2, rb2 - (rb + 2));
        consumed = rb2 - start + 1; return true;
    }

    private static bool TryParseCollapsedRef(string text, int start, MarkdownReaderOptions? options, out int consumed, out string label, InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches = null) {
        consumed = 0; label = string.Empty;
        if (start >= text.Length || text[start] != '[') return false;
        int rb = FindMatchingBracket(text, start, options, inlineHtmlWrapperMatches); if (rb < 0) return false;
        if (rb + 2 >= text.Length || text[rb + 1] != '[' || text[rb + 2] != ']') return false;
        label = text.Substring(start + 1, rb - (start + 1));
        consumed = rb + 3 - start;
        return true;
    }

    private static bool TryParseShortcutRef(string text, int start, MarkdownReaderOptions? options, out int consumed, out string label, InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches = null) {
        consumed = 0; label = string.Empty;
        if (start >= text.Length || text[start] != '[') return false;
        int rb = FindMatchingBracket(text, start, options, inlineHtmlWrapperMatches); if (rb < 0) return false;
        if (rb + 1 < text.Length && text[rb + 1] == '[') return false;
        label = text.Substring(start + 1, rb - (start + 1));
        consumed = rb + 1 - start;
        return true;
    }

    private static bool ContainsResolvedLinkInLabel(string label, MarkdownReaderOptions? options, MarkdownReaderState? state) {
        if (string.IsNullOrEmpty(label)) return false;
        var inlineHtmlWrapperMatches = BuildInlineHtmlWrapperMatchIndex(label);

        for (int i = 0; i < label.Length; i++) {
            if (TrySkipLinkLabelInlineSpan(label, i, options, out int spanConsumed, inlineHtmlWrapperMatches)) {
                i += spanConsumed - 1;
                continue;
            }

            if (label[i] != '[') {
                continue;
            }

            if (i > 0 && label[i - 1] == '!') {
                continue;
            }

            if (TryParseLink(
                label,
                i,
                options,
                sourceMap: null,
                state,
                out _,
                out _,
                out _,
                out _,
                out _,
                out _,
                out _,
                out _,
                inlineHtmlWrapperMatches)) {
                return true;
            }

            if (state == null) {
                continue;
            }

            if (TryParseRefLink(label, i, options, out _, out _, out var referenceLabel, inlineHtmlWrapperMatches)
                && state.LinkRefs.ContainsKey(NormalizeReferenceLabel(referenceLabel))) {
                return true;
            }

            if (TryParseCollapsedRef(label, i, options, out _, out var collapsedLabel, inlineHtmlWrapperMatches)
                && state.LinkRefs.ContainsKey(NormalizeReferenceLabel(collapsedLabel))) {
                return true;
            }

            if (TryParseShortcutRef(label, i, options, out _, out var shortcutLabel, inlineHtmlWrapperMatches)
                && state.LinkRefs.ContainsKey(NormalizeReferenceLabel(shortcutLabel))) {
                return true;
            }
        }

        return false;
    }

    private static bool TryParseLink(
        string text,
        int start,
        MarkdownReaderOptions? options,
        MarkdownInlineSourceMap? sourceMap,
        MarkdownReaderState? state,
        out int consumed,
        out string label,
        out string href,
        out string? title,
        out int hrefStart,
        out int hrefLength,
        out int? titleStart,
        out int? titleLength,
        InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches = null) {
        consumed = 0; label = href = string.Empty; title = null;
        hrefStart = 0; hrefLength = 0; titleStart = null; titleLength = null;
        if (start >= text.Length || text[start] != '[') return false;
        int labelEnd = FindMatchingBracket(text, start, options, inlineHtmlWrapperMatches);
        if (labelEnd < 0) return false;
        int parenOpen = (labelEnd + 1 < text.Length && text[labelEnd + 1] == '(') ? labelEnd + 1 : -1;
        if (parenOpen < 0) return false;
        int parenClose = FindMatchingParen(text, parenOpen);
        if (parenClose < 0) return false;
        label = text.Substring(start + 1, labelEnd - (start + 1));
        if (ContainsResolvedLinkInLabel(label, options, state)) return false;

        string inner = text.Substring(parenOpen + 1, parenClose - (parenOpen + 1));
        if (!TrySplitUrlAndOptionalTitle(
            inner,
            out href,
            out title,
            out int hrefInnerStart,
            out int hrefInnerLength,
            out int? titleInnerStart,
            out int? titleInnerLength,
            sourceMap,
            parenOpen + 1)) {
            if (!TryParseTrimmedLiteralDestination(inner, out href, out hrefInnerStart, out hrefInnerLength)) return false;
            title = null;
            titleInnerStart = null;
            titleInnerLength = null;
        }

        hrefStart = parenOpen + 1 + hrefInnerStart;
        hrefLength = hrefInnerLength;
        if (titleInnerStart.HasValue && titleInnerLength.HasValue) {
            titleStart = parenOpen + 1 + titleInnerStart.Value;
            titleLength = titleInnerLength.Value;
        }

        consumed = parenClose - start + 1;
        return true;
    }

    private static bool TrySplitUrlAndOptionalTitle(
        string? inner,
        out string url,
        out string? title,
        out int urlStart,
        out int urlLength,
        out int? titleStart,
        out int? titleLength,
        MarkdownInlineSourceMap? sourceMap = null,
        int sourceOffset = 0) {
        url = string.Empty;
        title = null;
        urlStart = 0;
        urlLength = 0;
        titleStart = null;
        titleLength = null;
        if (inner == null) return false;

        int start = 0;
        while (start < inner.Length && IsLinkWhitespace(inner[start])) {
            start++;
        }

        int endExclusive = inner.Length;
        while (endExclusive > start && IsLinkWhitespace(inner[endExclusive - 1])) {
            endExclusive--;
        }

        if (endExclusive <= start) {
            urlStart = start;
            urlLength = 0;
            url = string.Empty;
            title = null;
            return true;
        }

        // CommonMark: destination can be wrapped in <...> to allow spaces and parentheses safely.
        if (inner[start] == '<') {
            int gt = inner.IndexOf('>', start + 1);
            if (gt >= start + 1 && gt < endExclusive) {
                if (!IsValidAngleLinkDestination(inner, start + 1, gt)
                    || ContainsSourceLineBreak(sourceMap, sourceOffset + start + 1, gt - (start + 1))) {
                    return false;
                }

                urlStart = start + 1;
                urlLength = gt - urlStart;
                url = DecodeLinkDestination(inner.Substring(urlStart, urlLength));

                int restStart = gt + 1;
                if (restStart < endExclusive && !IsLinkWhitespace(inner[restStart])) {
                    return false;
                }

                while (restStart < endExclusive && IsLinkWhitespace(inner[restStart])) {
                    restStart++;
                }

                if (restStart >= endExclusive) {
                    return true;
                }

                if (!TryParseOptionalTitleToken(inner, restStart, endExclusive, out title, out int parsedTitleStart, out int parsedTitleLength)) {
                    return false;
                }

                title = DecodeLinkDestinationOrTitle(title!);
                titleStart = parsedTitleStart;
                titleLength = parsedTitleLength;
                return true;
            }
        }

        int ws = -1;
        for (int i = start; i < endExclusive; i++) {
            if (IsLinkWhitespace(inner[i])) {
                ws = i;
                break;
            }
        }

        if (ws < 0) {
            urlStart = start;
            urlLength = endExclusive - start;
            url = DecodeLinkDestination(inner.Substring(urlStart, urlLength));
            title = null;
            return true;
        }

        urlStart = start;
        urlLength = ws - start;
        url = DecodeLinkDestination(inner.Substring(urlStart, urlLength));

        int remainingStart = ws;
        while (remainingStart < endExclusive && IsLinkWhitespace(inner[remainingStart])) {
            remainingStart++;
        }

        if (remainingStart >= endExclusive) { title = null; return true; }

        if (!TryParseOptionalTitleToken(inner, remainingStart, endExclusive, out title, out int parsedStart, out int parsedLength)) return false;
        title = DecodeLinkDestinationOrTitle(title!);
        titleStart = parsedStart;
        titleLength = parsedLength;
        return true;
    }

    private static bool TrySplitUrlAndOptionalTitle(string? inner, out string url, out string? title) =>
        TrySplitUrlAndOptionalTitle(inner, out url, out title, out _, out _, out _, out _);

    private static bool ContainsSourceLineBreak(MarkdownInlineSourceMap? sourceMap, int startIndex, int length) =>
        sourceMap != null && sourceMap.ContainsSourceLineBreak(startIndex, length);

    private static bool TryParseTrimmedLiteralDestination(string inner, out string destination, out int destinationStart, out int destinationLength) {
        destination = string.Empty;
        destinationStart = 0;
        destinationLength = 0;

        int trimmedStart = 0;
        while (trimmedStart < inner.Length && IsLinkWhitespace(inner[trimmedStart])) {
            trimmedStart++;
        }

        int trimmedEndExclusive = inner.Length;
        while (trimmedEndExclusive > trimmedStart && IsLinkWhitespace(inner[trimmedEndExclusive - 1])) {
            trimmedEndExclusive--;
        }

        if (trimmedEndExclusive <= trimmedStart) return false;
        if (inner[trimmedStart] == '<') return false;
        if (IndexOfWhitespace(inner, trimmedStart, trimmedEndExclusive) >= 0) return false;

        destination = DecodeLinkDestination(inner.Substring(trimmedStart, trimmedEndExclusive - trimmedStart));
        destinationStart = trimmedStart;
        destinationLength = Math.Max(0, trimmedEndExclusive - trimmedStart);
        return true;
    }

    private static int IndexOfWhitespace(string s) {
        for (int i = 0; i < s.Length; i++) if (IsLinkWhitespace(s[i])) return i;
        return -1;
    }

    private static int IndexOfWhitespace(string s, int start, int endExclusive) {
        for (int i = start; i < endExclusive; i++) if (IsLinkWhitespace(s[i])) return i;
        return -1;
    }

    private static bool IsLinkWhitespace(char c) =>
        c == ' ' || c == '\t' || c == '\n' || c == '\r';

    private static bool IsValidAngleLinkDestination(string value, int start, int endExclusive) {
        for (int i = start; i < endExclusive; i++) {
            char c = value[i];
            if (c == '\n' || c == '\r' || c == '<') {
                return false;
            }
        }

        return true;
    }

    private static string? TryParseOptionalTitleToken(string s) {
        if (string.IsNullOrWhiteSpace(s)) return null;
        int start = 0;
        while (start < s.Length && char.IsWhiteSpace(s[start])) {
            start++;
        }

        int endExclusive = s.Length;
        while (endExclusive > start && char.IsWhiteSpace(s[endExclusive - 1])) {
            endExclusive--;
        }

        return TryParseOptionalTitleToken(s, start, endExclusive, out string? title, out _, out _) ? title : null;
    }

    private static bool TryParseOptionalTitleToken(
        string s,
        int start,
        int endExclusive,
        out string? title,
        out int titleStart,
        out int titleLength) {
        title = null;
        titleStart = 0;
        titleLength = 0;
        if (string.IsNullOrEmpty(s) || endExclusive - start < 2) {
            return false;
        }

        char open = s[start];
        char close = s[endExclusive - 1];
        if ((open == '"' && close == '"') ||
            (open == '\'' && close == '\'') ||
            (open == '(' && close == ')')) {
            titleStart = start + 1;
            titleLength = endExclusive - start - 2;
            if (ContainsUnescapedTitleDelimiter(s, titleStart, titleStart + titleLength, close)) {
                return false;
            }

            title = s.Substring(titleStart, titleLength);
            return true;
        }

        return false;
    }

    private static int FindMatchingParen(string text, int openIndex) {
        int depth = 0;
        bool inDoubleQuotes = false;
        bool inSingleQuotes = false;
        bool inAngle = false;
        bool escaped = false;
        for (int i = openIndex; i < text.Length; i++) {
            char c = text[i];
            if (escaped) {
                escaped = false;
                continue;
            }
            if (c == '\\') {
                escaped = true;
                continue;
            }
            if (inAngle) {
                if (c == '>') inAngle = false;
                continue;
            }
            if (inDoubleQuotes) {
                if (c == '"') inDoubleQuotes = false;
                continue;
            }
            if (inSingleQuotes) {
                if (c == '\'') inSingleQuotes = false;
                continue;
            }
            if (c == '(') { depth++; continue; }
            if (c == ')') { depth--; if (depth == 0) return i; continue; }
            if (depth == 1) {
                if (c == '<') { inAngle = true; continue; }
                if (c == '"') { inDoubleQuotes = true; continue; }
                if (c == '\'') { inSingleQuotes = true; continue; }
            }
        }
        return -1;
    }

    private static bool ContainsUnescapedTitleDelimiter(string value, int start, int endExclusive, char delimiter) {
        bool escaped = false;
        for (int i = start; i < endExclusive; i++) {
            char c = value[i];
            if (escaped) {
                escaped = false;
                continue;
            }

            if (c == '\\') {
                escaped = true;
                continue;
            }

            if (c == delimiter) {
                return true;
            }
        }

        return false;
    }
}

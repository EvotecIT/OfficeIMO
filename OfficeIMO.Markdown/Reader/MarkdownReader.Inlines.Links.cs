namespace OfficeIMO.Markdown;

internal static partial class MarkdownReader {
    private static int FindMatchingBracket(
        string text,
        int openIndex,
        MarkdownLoadOptions? options = null,
        InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches = null) {
        if (string.IsNullOrEmpty(text) || openIndex < 0 || openIndex >= text.Length || text[openIndex] != '[') return -1;

        var index = InlineScanIndexCache.GetBracketIndex(text, options, inlineHtmlWrapperMatches);
        if (index != null && index.TryGetMatch(openIndex, out int closeIndex)) {
            return closeIndex;
        }

        return ScanForMatchingBracket(text, openIndex, text.Length, options, inlineHtmlWrapperMatches);
    }

    // Matches a bracket as if the text ended at end. Spans that cross end are the only way a scan of
    // the whole text can differ from a scan of the shorter text, and taking one moves the scan past
    // end, so a match found before end is the same in both; anything else is rescanned within end.
    private static int FindMatchingBracketWithin(
        string text,
        int openIndex,
        int end,
        MarkdownLoadOptions? options,
        InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches) {
        if (end >= text.Length) return FindMatchingBracket(text, openIndex, options, inlineHtmlWrapperMatches);
        if (openIndex < 0 || openIndex >= end || text[openIndex] != '[') return -1;

        int closeIndex = FindMatchingBracket(text, openIndex, options, inlineHtmlWrapperMatches);
        return closeIndex >= 0 && closeIndex < end
            ? closeIndex
            : ScanForMatchingBracket(text, openIndex, end, options, inlineHtmlWrapperMatches);
    }

    private static int ScanForMatchingBracket(
        string text,
        int openIndex,
        int end,
        MarkdownLoadOptions? options,
        InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches) {
        int depth = 0;
        bool escaped = false;
        var effectiveInlineHtmlWrapperMatches = inlineHtmlWrapperMatches;
        for (int i = openIndex; i < end; i++) {
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

            if (TrySkipLinkLabelInlineSpan(text, i, end, options, out int spanConsumed, effectiveInlineHtmlWrapperMatches)) {
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
        MarkdownLoadOptions? options,
        out int consumed,
        InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches = null) =>
        TrySkipLinkLabelInlineSpan(text, start, text?.Length ?? 0, options, out consumed, inlineHtmlWrapperMatches);

    // A span counts only when it ends by end, which is what parsing the text cut at end finds: each
    // span closes at the first terminator it reaches, and end is never inside a backtick run.
    private static bool TrySkipLinkLabelInlineSpan(
        string text,
        int start,
        int end,
        MarkdownLoadOptions? options,
        out int consumed,
        InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches = null) {
        consumed = 0;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= end) return false;

        if (text[start] == '`' && TryConsumeMatchedBacktickSpan(text, start, out consumed) && start + consumed <= end) {
            return true;
        }

        if (text[start] == '<') {
            bool inlineHtml = options?.InlineHtml != false;
            if (inlineHtml && TryConsumeSupportedInlineHtmlWrapperSpan(text, start, out consumed, inlineHtmlWrapperMatches) && start + consumed <= end) {
                return true;
            }

            if (TryParseAngleAutolink(text, start, out consumed, out _, out _) && start + consumed <= end) {
                return true;
            }

            if (inlineHtml && TryConsumeRawInlineHtmlTag(text, start, out consumed) && start + consumed <= end) {
                return true;
            }
        }

        consumed = 0;
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

        int closingStart = FindClosingBacktickRun(text, start + fenceLength, fenceLength);
        if (closingStart < 0) {
            return false;
        }

        consumed = closingStart + fenceLength - start;
        return true;
    }

    // Finds the next run of exactly fenceLength backticks at or after from, which must not be inside a run.
    private static int FindClosingBacktickRun(string text, int from, int fenceLength) {
        var index = InlineScanIndexCache.GetBacktickIndex(text);
        if (index != null) {
            return index.FindRun(from, fenceLength);
        }

        int scan = from;
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
                return scan;
            }

            scan += runLength;
        }

        return -1;
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

    // CommonMark caps link labels at 999 characters. Longer bracketed text can never name a
    // definition, so it is not normalized and looked up for every bracket it contains.
    private const int MaxReferenceLabelLength = 999;

    // The label isn't capped like the reference, so its length is returned rather than a copy.
    private static bool TryParseRefLink(string text, int start, MarkdownLoadOptions? options, out int consumed, out int labelLength, out string refLabel, InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches = null) {
        consumed = 0; labelLength = 0; refLabel = string.Empty;
        if (start >= text.Length || text[start] != '[') return false;
        int rb = FindMatchingBracket(text, start, options, inlineHtmlWrapperMatches); if (rb < 0) return false;
        if (rb + 1 >= text.Length || text[rb + 1] != '[') return false;
        int rb2 = FindMatchingBracket(text, rb + 1, options, inlineHtmlWrapperMatches); if (rb2 < 0) return false;
        if (rb2 - rb - 2 > MaxReferenceLabelLength) return false;
        labelLength = rb - (start + 1);
        refLabel = text.Substring(rb + 2, rb2 - (rb + 2));
        consumed = rb2 - start + 1; return true;
    }

    private static bool TryParseCollapsedRef(string text, int start, MarkdownLoadOptions? options, out int consumed, out string label, InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches = null) {
        consumed = 0; label = string.Empty;
        if (start >= text.Length || text[start] != '[') return false;
        int rb = FindMatchingBracket(text, start, options, inlineHtmlWrapperMatches); if (rb < 0) return false;
        if (rb + 2 >= text.Length || text[rb + 1] != '[' || text[rb + 2] != ']') return false;
        if (rb - start - 1 > MaxReferenceLabelLength) return false;
        label = text.Substring(start + 1, rb - (start + 1));
        consumed = rb + 3 - start;
        return true;
    }

    private static bool TryParseShortcutRef(string text, int start, MarkdownLoadOptions? options, out int consumed, out string label, InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches = null) {
        consumed = 0; label = string.Empty;
        if (start >= text.Length || text[start] != '[') return false;
        int rb = FindMatchingBracket(text, start, options, inlineHtmlWrapperMatches); if (rb < 0) return false;
        if (rb + 1 < text.Length && text[rb + 1] == '[') return false;
        if (rb - start - 1 > MaxReferenceLabelLength) return false;
        label = text.Substring(start + 1, rb - (start + 1));
        consumed = rb + 1 - start;
        return true;
    }

    // Checks the label text[start..end) of a bracket pair as if it were a text of its own, so every
    // construct found in it has to end inside it. Copying and indexing each label separately made
    // nested brackets quadratic, because every opener's label holds all the inner ones.
    private static bool ContainsResolvedLinkInLabel(
        string text, int start, int end, MarkdownLoadOptions? options, MarkdownReaderState? state,
        InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches) {
        if (end <= start) return false;
        // Wrapper pairs inside the label pair the same way in an index of the whole text; a pair
        // that crosses the label's bounds is rejected by the bounded checks below. Image brackets
        // in a label are matched around wrappers even when the parse has inline HTML turned off.
        var matches = inlineHtmlWrapperMatches != null && options?.InlineHtml != false
            ? inlineHtmlWrapperMatches
            : InlineScanIndexCache.GetWrapperIndex(text);

        for (int i = start; i < end; i++) {
            char c = text[i];
            if (c == '\\' && i + 1 < end && IsBackslashEscapable(text[i + 1])) {
                i++;
                continue;
            }
            if (c == '!' && TrySkipLinkLabelImage(text, i, end, state, matches, out int imageLength)) {
                i += imageLength - 1;
                continue;
            }
            if (TrySkipLinkLabelInlineSpan(text, i, end, options, out int spanConsumed, matches)) {
                if (c == '<' && TryParseAngleAutolink(text, i, out int angleLength, out _, out _) && i + angleLength <= end) return true;
                if (c == '<' && matches.TryGet(i, out string tagName, out int closingStart) && closingStart + tagName.Length + 3 <= end &&
                    ContainsAngleAutolinkInLabelSpan(text, i + tagName.Length + 2, closingStart, end, options, state, matches)) return true;
                i += spanConsumed - 1;
                continue;
            }

            if (c != '[') {
                continue;
            }

            if (i > start && text[i - 1] == '!') {
                continue;
            }

            if (TryMatchInlineLink(text, i, end, options, matches, out _, out int parenOpen, out int parenClose)
                && TryParseInlineLinkDestination(text, parenOpen, parenClose, sourceMap: null, out _, out _, out _, out _, out _, out _)) {
                return true;
            }

            if (state == null) {
                continue;
            }

            int labelEnd = FindMatchingBracketWithin(text, i, end, options, matches);
            if (labelEnd < 0) {
                continue;
            }

            if (labelEnd + 1 < end && text[labelEnd + 1] == '[') {
                // Full reference [label][ref]
                int refEnd = FindMatchingBracketWithin(text, labelEnd + 1, end, options, matches);
                if (refEnd >= 0 && refEnd - labelEnd - 2 <= MaxReferenceLabelLength
                    && state.LinkRefs.ContainsKey(NormalizeReferenceLabel(text.Substring(labelEnd + 2, refEnd - labelEnd - 2)))) {
                    return true;
                }

                // Collapsed reference [label][]
                if (labelEnd + 2 < end && text[labelEnd + 2] == ']' && labelEnd - i - 1 <= MaxReferenceLabelLength
                    && state.LinkRefs.ContainsKey(NormalizeReferenceLabel(text.Substring(i + 1, labelEnd - i - 1)))) {
                    return true;
                }

                continue;
            }

            // Shortcut reference [label]
            if (labelEnd - i - 1 <= MaxReferenceLabelLength
                && state.LinkRefs.ContainsKey(NormalizeReferenceLabel(text.Substring(i + 1, labelEnd - i - 1)))) {
                return true;
            }
        }

        return false;
    }

    // labelEnd bounds the images, which may extend past the wrapper before being rejected.
    private static bool ContainsAngleAutolinkInLabelSpan(
        string text, int start, int end, int labelEnd, MarkdownLoadOptions? options, MarkdownReaderState? state,
        InlineHtmlWrapperMatchIndex matches) {
        // Inspect wrapper contents without recursively rescanning nested wrappers. Code, escapes,
        // HTML attributes and image alternatives are literal contexts rather than nested links.
        // Index the next equal-length backtick run once: probing every possible opener
        // separately rescans unmatched suffixes quadratically.
        var nextBacktickRun = new Dictionary<int, int>();
        var codeSpanEnds = new Dictionary<int, int>();
        for (int runEnd = end; runEnd > start;) {
            int runStart = runEnd - 1;
            if (text[runStart] != '`') { runEnd--; continue; }
            while (runStart > start && text[runStart - 1] == '`') runStart--;
            int length = runEnd - runStart;
            if (nextBacktickRun.TryGetValue(length, out int closer))
                codeSpanEnds[runStart] = closer + length;
            nextBacktickRun[length] = runStart;
            runEnd = runStart;
        }
        for (int i = start; i < end; i++) {
            if (text[i] == '\\' && i + 1 < end && IsBackslashEscapable(text[i + 1])) { i++; continue; }
            if (text[i] == '`') {
                int runStart = i;
                while (i + 1 < end && text[i + 1] == '`') i++;
                if (codeSpanEnds.TryGetValue(runStart, out int codeSpanEnd)) i = codeSpanEnd - 1;
                continue;
            }
            if (text[i] == '!' && TrySkipLinkLabelImage(text, i, labelEnd, state, matches, out int imageLength) && i + imageLength <= end) {
                i += imageLength - 1;
                continue;
            }
            if (text[i] != '<') continue;
            if (TryParseAngleAutolink(text, i, out int angleLength, out _, out _) && i + angleLength <= end) return true;
            if (options?.InlineHtml != false && TryConsumeRawInlineHtmlTag(text, i, out int htmlLength) && i + htmlLength <= end) i += htmlLength - 1;
        }
        return false;
    }

    // An inline image ![alt](src), or with definitions a reference image, that ends by end.
    private static bool TrySkipLinkLabelImage(string text, int start, int end, MarkdownReaderState? state, InlineHtmlWrapperMatchIndex matches, out int consumed) {
        consumed = 0;
        if (start + 1 >= end || text[start] != '!' || text[start + 1] != '[') return false;
        int altEnd = FindMatchingBracketWithin(text, start + 1, end, options: null, matches);
        if (altEnd < 0) return false;

        if (altEnd + 1 < end && text[altEnd + 1] == '(') {
            int parenClose = FindMatchingParen(text, altEnd + 1, end);
            if (parenClose >= 0 && TryParseInlineLinkDestination(text, altEnd + 1, parenClose, sourceMap: null, out _, out _, out _, out _, out _, out _)) {
                consumed = parenClose - start + 1;
                return true;
            }

            // A reference image can't be followed by "(".
            return false;
        }

        if (state == null) return false;
        if (altEnd + 1 < end && text[altEnd + 1] == '[') {
            int labelEnd = FindMatchingBracketWithin(text, altEnd + 1, end, options: null, matches);
            if (labelEnd < 0) return false;
            consumed = labelEnd - start + 1;
            return true;
        }

        consumed = altEnd - start + 1;
        return true;
    }

    // The bracket and parenthesis positions of an inline link [label](...) that ends by end.
    private static bool TryMatchInlineLink(
        string text, int start, int end, MarkdownLoadOptions? options, InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches,
        out int labelEnd, out int parenOpen, out int parenClose) {
        parenOpen = parenClose = -1;
        labelEnd = FindMatchingBracketWithin(text, start, end, options, inlineHtmlWrapperMatches);
        if (labelEnd < 0) return false;
        if (labelEnd + 1 >= end || text[labelEnd + 1] != '(') return false;
        parenOpen = labelEnd + 1;
        parenClose = FindMatchingParen(text, parenOpen, end);
        return parenClose >= 0;
    }

    // Parses the destination and optional title between a link's parentheses.
    private static bool TryParseInlineLinkDestination(
        string text, int parenOpen, int parenClose, MarkdownInlineSourceMap? sourceMap,
        out string href, out string? title, out int hrefStart, out int hrefLength, out int? titleStart, out int? titleLength) {
        hrefStart = hrefLength = 0;
        titleStart = titleLength = null;
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

        return true;
    }

    private static bool TryParseLink(
        string text,
        int start,
        MarkdownLoadOptions? options,
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
        InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches = null,
        bool checkNestedLinks = true) {
        consumed = 0; label = href = string.Empty; title = null;
        hrefStart = 0; hrefLength = 0; titleStart = null; titleLength = null;
        if (start >= text.Length || text[start] != '[') return false;
        if (!TryMatchInlineLink(text, start, text.Length, options, inlineHtmlWrapperMatches, out int labelEnd, out int parenOpen, out int parenClose)) return false;
        // The label scan visits every inner bracket, so an inner candidate needs no scan of its own:
        // any link it contains is also found directly. Rescanning would be exponential in the depth.
        if (checkNestedLinks && ContainsResolvedLinkInLabel(text, start + 1, labelEnd, options, state, inlineHtmlWrapperMatches)) return false;
        if (!TryParseInlineLinkDestination(text, parenOpen, parenClose, sourceMap, out href, out title, out hrefStart, out hrefLength, out titleStart, out titleLength)) return false;

        label = text.Substring(start + 1, labelEnd - (start + 1));
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

    // The opening parenthesis counts as the first level.
    private const int MaxLinkDestinationParenDepth = 33;

    private static int FindMatchingParen(string text, int openIndex) => FindMatchingParen(text, openIndex, text.Length);

    private static int FindMatchingParen(string text, int openIndex, int end) {
        int depth = 0;
        bool inDoubleQuotes = false;
        bool inSingleQuotes = false;
        bool inAngle = false;
        bool escaped = false;
        for (int i = openIndex; i < end; i++) {
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
                // A pointed destination can't contain '<' or a line ending; stopping here keeps an
                // unclosed '<' from scanning the rest of the text for every link opener.
                if (c == '<' || c == '\n') return -1;
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
            if (c == '(') {
                // CommonMark lets implementations cap parenthesis nesting in destinations; like cmark,
                // give up past 32 so unbalanced input can't make every opener scan to the end.
                if (++depth > MaxLinkDestinationParenDepth) return -1;
                continue;
            }
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

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private const int MaxImageAltNestingDepth = 32;

    private static string ExtractImageAltPlainText(
        string altMarkdown,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        int imageAltDepth = 0) {
        if (string.IsNullOrEmpty(altMarkdown)) {
            return string.Empty;
        }

        if (imageAltDepth >= MaxImageAltNestingDepth) {
            return altMarkdown;
        }

        var altSequence = ParseInlinesInternal(
            altMarkdown,
            options,
            state,
            allowLinks: true,
            allowImages: true,
            imageAltDepth: imageAltDepth + 1);
        return InlinePlainText.Extract(altSequence);
    }

    private static bool TryConsumeLiteralInlineImage(string text, int start, out int consumed, InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches = null) {
        consumed = 0;
        if (start + 1 >= text.Length || text[start] != '!' || text[start + 1] != '[') return false;
        int altEnd = FindMatchingBracket(text, start + 1, inlineHtmlWrapperMatches: inlineHtmlWrapperMatches);
        if (altEnd < 0) return false;
        if (altEnd + 1 >= text.Length || text[altEnd + 1] != '(') return false;
        int parenClose = FindMatchingParen(text, altEnd + 1);
        if (parenClose < 0) return false;
        consumed = parenClose - start + 1;
        return true;
    }

    private static bool TryParseImageLink(string text, int start, out int consumed, out string alt, out string img, out string? imgTitle, out string href, out string? hrefTitle) =>
        TryParseImageLink(text, start, null, out consumed, out alt, out img, out imgTitle, out href, out hrefTitle, out _, out _, out _, out _, out _, out _, out _, out _, out _, out _);

    private static bool TryParseImageLink(
        string text,
        int start,
        MarkdownInlineSourceMap? sourceMap,
        out int consumed,
        out string alt,
        out string img,
        out string? imgTitle,
        out string href,
        out string? hrefTitle,
        out int altStart,
        out int altLength,
        out int imgStart,
        out int imgLength,
        out int? imgTitleStart,
        out int? imgTitleLength,
        out int hrefStart,
        out int hrefLength,
        out int? hrefTitleStart,
        out int? hrefTitleLength,
        InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches = null) {
        consumed = 0; alt = img = href = string.Empty; imgTitle = hrefTitle = null;
        altStart = altLength = imgStart = imgLength = hrefStart = hrefLength = 0;
        imgTitleStart = imgTitleLength = hrefTitleStart = hrefTitleLength = null;
        if (start >= text.Length || text[start] != '[') return false;
        if (start + 1 >= text.Length || text[start + 1] != '!') return false;
        if (start + 2 >= text.Length || text[start + 2] != '[') return false;
        int altEnd = FindMatchingBracket(text, start + 2, inlineHtmlWrapperMatches: inlineHtmlWrapperMatches);
        if (altEnd < 0) return false;
        if (altEnd + 1 >= text.Length || text[altEnd + 1] != '(') return false;
        int imgClose = FindMatchingParen(text, altEnd + 1);
        if (imgClose < 0) return false;
        altStart = start + 3;
        altLength = altEnd - altStart;
        alt = text.Substring(altStart, altLength);
        string inner = text.Substring(altEnd + 2, imgClose - (altEnd + 2));
        if (!TrySplitUrlAndOptionalTitle(
            inner,
            out img,
            out imgTitle,
            out int imgInnerStart,
            out int imgInnerLength,
            out int? imgTitleInnerStart,
            out int? imgTitleInnerLength,
            sourceMap,
            altEnd + 2)) {
            if (!TryParseTrimmedLiteralDestination(inner, out img, out imgInnerStart, out imgInnerLength)) return false;
            imgTitle = null;
            imgTitleInnerStart = null;
            imgTitleInnerLength = null;
        }
        imgStart = altEnd + 2 + imgInnerStart;
        imgLength = imgInnerLength;
        imgTitleStart = imgTitleInnerStart.HasValue ? altEnd + 2 + imgTitleInnerStart.Value : null;
        imgTitleLength = imgTitleInnerLength;
        int closeBracket = (imgClose + 1 < text.Length && text[imgClose + 1] == ']') ? imgClose + 1 : -1;
        if (closeBracket < 0) return false;
        int parenOpen2 = (closeBracket + 1 < text.Length && text[closeBracket + 1] == '(') ? closeBracket + 1 : -1;
        if (parenOpen2 != closeBracket + 1) return false;
        int parenClose2 = FindMatchingParen(text, parenOpen2);
        if (parenClose2 < 0) return false;
        string hrefInner = text.Substring(parenOpen2 + 1, parenClose2 - (parenOpen2 + 1));
        if (!TrySplitUrlAndOptionalTitle(
            hrefInner,
            out href,
            out hrefTitle,
            out int hrefInnerStart,
            out int hrefInnerLength,
            out int? hrefTitleInnerStart,
            out int? hrefTitleInnerLength,
            sourceMap,
            parenOpen2 + 1)) {
            if (!TryParseTrimmedLiteralDestination(hrefInner, out href, out hrefInnerStart, out hrefInnerLength)) return false;
            hrefTitle = null;
            hrefTitleInnerStart = null;
            hrefTitleInnerLength = null;
        }
        hrefStart = parenOpen2 + 1 + hrefInnerStart;
        hrefLength = hrefInnerLength;
        hrefTitleStart = hrefTitleInnerStart.HasValue ? parenOpen2 + 1 + hrefTitleInnerStart.Value : null;
        hrefTitleLength = hrefTitleInnerLength;
        consumed = parenClose2 - start + 1;
        return true;
    }

    private static bool TryParseInlineImage(string text, int start, out int consumed, out string alt, out string src, out string? title) =>
        TryParseInlineImage(
            text,
            start,
            out consumed,
            out alt,
            out src,
            out title,
            out _,
            out _,
            out _,
            out _,
            out _,
            out _);

    private static bool TryParseInlineImage(
        string text,
        int start,
        out int consumed,
        out string alt,
        out string src,
        out string? title,
        out int altStart,
        out int altLength,
        out int srcStart,
        out int srcLength,
        out int? titleStart,
        out int? titleLength,
        InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches = null) {
        return TryParseInlineImage(
            text,
            start,
            null,
            out consumed,
            out alt,
            out src,
            out title,
            out altStart,
            out altLength,
            out srcStart,
            out srcLength,
            out titleStart,
            out titleLength,
            inlineHtmlWrapperMatches);
    }

    private static bool TryParseInlineImage(
        string text,
        int start,
        MarkdownInlineSourceMap? sourceMap,
        out int consumed,
        out string alt,
        out string src,
        out string? title,
        out int altStart,
        out int altLength,
        out int srcStart,
        out int srcLength,
        out int? titleStart,
        out int? titleLength,
        InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches = null) {
        consumed = 0;
        alt = src = string.Empty;
        title = null;
        altStart = altLength = srcStart = srcLength = 0;
        titleStart = titleLength = null;
        if (start + 1 >= text.Length || text[start] != '!' || text[start + 1] != '[') return false;
        int altEnd = FindMatchingBracket(text, start + 1, inlineHtmlWrapperMatches: inlineHtmlWrapperMatches);
        if (altEnd < 0) return false;
        if (altEnd + 1 >= text.Length || text[altEnd + 1] != '(') return false;
        int parenClose = FindMatchingParen(text, altEnd + 1);
        if (parenClose < 0) return false;
        altStart = start + 2;
        altLength = altEnd - altStart;
        alt = text.Substring(altStart, altLength);
        string inner = text.Substring(altEnd + 2, parenClose - (altEnd + 2));
        if (!TrySplitUrlAndOptionalTitle(
            inner,
            out src,
            out title,
            out int srcInnerStart,
            out int srcInnerLength,
            out int? titleInnerStart,
            out int? titleInnerLength,
            sourceMap,
            altEnd + 2)) {
            if (!TryParseTrimmedLiteralDestination(inner, out src, out srcInnerStart, out srcInnerLength)) return false;
            title = null;
            titleInnerStart = null;
            titleInnerLength = null;
        }

        srcStart = altEnd + 2 + srcInnerStart;
        srcLength = srcInnerLength;
        titleStart = titleInnerStart.HasValue ? altEnd + 2 + titleInnerStart.Value : null;
        titleLength = titleInnerLength;
        consumed = parenClose - start + 1;
        return true;
    }

    private static bool TryParseReferenceImage(string text, int start, out int consumed, out string alt, out string label) =>
        TryParseReferenceImage(text, start, out consumed, out alt, out label, out _, out _);

    private static bool TryParseReferenceImage(
        string text,
        int start,
        out int consumed,
        out string alt,
        out string label,
        out int altStart,
        out int altLength,
        InlineHtmlWrapperMatchIndex? inlineHtmlWrapperMatches = null) {
        consumed = 0;
        alt = label = string.Empty;
        altStart = altLength = 0;
        if (start + 1 >= text.Length || text[start] != '!' || text[start + 1] != '[') return false;
        int altEnd = FindMatchingBracket(text, start + 1, inlineHtmlWrapperMatches: inlineHtmlWrapperMatches);
        if (altEnd < 0) return false;

        altStart = start + 2;
        altLength = altEnd - altStart;
        alt = text.Substring(altStart, altLength);

        // Inline image uses "(...)" and is handled elsewhere.
        if (altEnd + 1 < text.Length && text[altEnd + 1] == '(') return false;

        // Full or collapsed reference: ![alt][label] or ![alt][]
        if (altEnd + 1 < text.Length && text[altEnd + 1] == '[') {
            int labelEnd = FindMatchingBracket(text, altEnd + 1, inlineHtmlWrapperMatches: inlineHtmlWrapperMatches);
            if (labelEnd < 0) return false;
            label = text.Substring(altEnd + 2, labelEnd - (altEnd + 2));
            if (string.IsNullOrEmpty(label)) label = alt;
            consumed = labelEnd - start + 1;
            return true;
        }

        // Shortcut: ![label]
        label = alt;
        consumed = altEnd - start + 1;
        return true;
    }
}

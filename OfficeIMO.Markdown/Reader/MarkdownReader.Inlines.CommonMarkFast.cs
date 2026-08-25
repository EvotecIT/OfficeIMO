namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private enum SimpleInlineTokenKind {
        Strong,
        Emphasis,
        Code,
        Link,
        Image
    }

    private readonly record struct SimpleInlineToken(
        SimpleInlineTokenKind Kind,
        int Start,
        int End,
        int ContentStart,
        int ContentLength,
        int TargetStart = 0,
        int TargetLength = 0);

    private static bool TryParseSimpleCommonMarkInlines(
        string text,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        MarkdownInlineSourceMap? sourceMap,
        out InlineSequence sequence) {
        sequence = null!;
        if (state?.CaptureSyntaxTree != false
            || sourceMap != null
            || options.InlineParserExtensions.Count != 0
            || options.InlineTransformExtensions.Count != 0
            || (options.Abbreviations && state?.Abbreviations.Count > 0)
            || ContainsPotentialBareAutolinkSyntax(text, options)
            || string.IsNullOrEmpty(text)) {
            return false;
        }

        bool allowLiteralBrackets = !options.Footnotes && (state == null || state.LinkRefs.Count == 0);
        if (!ValidateSimpleCommonMarkInlineTokens(text, options, allowLiteralBrackets)) {
            return false;
        }

        sequence = new InlineSequence { AutoSpacing = false };
        int position = 0;
        int textStart = 0;
        while (position < text.Length) {
            if (!TryReadSimpleCommonMarkInlineToken(text, position, options, out var token)) {
                position++;
                continue;
            }

            AddSimpleTextRun(sequence, text, textStart, token.Start - textStart);
            switch (token.Kind) {
                case SimpleInlineTokenKind.Strong:
                    var strongContent = new InlineSequence { AutoSpacing = false };
                    AddSimpleEmphasisContent(strongContent, text, token.ContentStart, token.ContentLength, options);
                    sequence.AddRaw(new BoldSequenceInline(strongContent));
                    break;

                case SimpleInlineTokenKind.Emphasis:
                    var emphasisContent = new InlineSequence { AutoSpacing = false };
                    AddSimpleEmphasisContent(emphasisContent, text, token.ContentStart, token.ContentLength, options);
                    sequence.AddRaw(new ItalicSequenceInline(emphasisContent));
                    break;

                case SimpleInlineTokenKind.Code:
                    sequence.AddRaw(new CodeSpanInline(text.Substring(token.ContentStart, token.ContentLength)));
                    break;

                case SimpleInlineTokenKind.Link:
                    string linkContent = text.Substring(token.ContentStart, token.ContentLength);
                    string target = text.Substring(token.TargetStart, token.TargetLength);
                    string? resolvedTarget = string.IsNullOrWhiteSpace(target)
                        ? string.Empty
                        : ResolveUrl(target, options);
                    if (resolvedTarget == null) {
                        sequence.AddRaw(new MarkdownTextRun(linkContent));
                    } else {
                        var label = new InlineSequence { AutoSpacing = false };
                        label.AddRaw(new MarkdownTextRun(linkContent));
                        sequence.AddRaw(new LinkInline(label, resolvedTarget, title: null));
                    }
                    break;

                case SimpleInlineTokenKind.Image:
                    string altText = text.Substring(token.ContentStart, token.ContentLength);
                    string source = text.Substring(token.TargetStart, token.TargetLength);
                    string? resolvedSource = string.IsNullOrWhiteSpace(source)
                        ? string.Empty
                        : ResolveUrl(source, options);
                    if (resolvedSource == null) {
                        sequence.AddRaw(new MarkdownTextRun(altText));
                    } else {
                        sequence.AddRaw(new ImageInline(altText, resolvedSource));
                    }
                    break;
            }

            position = token.End;
            textStart = position;
        }

        AddSimpleTextRun(sequence, text, textStart, text.Length - textStart);
        return true;
    }

    private static bool ValidateSimpleCommonMarkInlineTokens(
        string text,
        MarkdownReaderOptions options,
        bool allowLiteralBrackets) {
        bool foundToken = false;
        bool foundLiteralBracket = false;
        int position = 0;
        while (position < text.Length) {
            char value = text[position];
            if (value is '*' or '_' or '`' or '[' or '!') {
                if (!TryReadSimpleCommonMarkInlineToken(text, position, options, out var token)) {
                    // Without footnotes or collected reference definitions, an unmatched '['
                    // cannot become a semantic link token. Keep scanning so ordinary bracketed
                    // prose (including task markers and literal callout labels) stays on the
                    // lightweight semantic path. A non-image '!' is likewise plain text.
                    if (allowLiteralBrackets
                        && value is '[' or '!'
                        && !IsPotentialInlineLinkOrImage(text, position)) {
                        foundLiteralBracket = true;
                        position++;
                        continue;
                    }

                    // A delimiter run that cannot open has no unresolved frame to affect on
                    // this validating path. It is literal CommonMark text even when it can
                    // close in isolation (for example, a trailing unmatched "**").
                    if (value is '*' or '_') {
                        int runLength = 1;
                        while (position + runLength < text.Length && text[position + runLength] == value) {
                            runLength++;
                        }

                        GetDelimiterFlags(
                            text,
                            position,
                            value,
                            runLength,
                            options.CjkFriendlyEmphasis,
                            out bool canOpen,
                            out _);
                        if (!canOpen) {
                            position += runLength;
                            continue;
                        }
                    }

                    return false;
                }

                foundToken = true;
                position = token.End;
                continue;
            }

            if (value is '\\' or '&' or '\n' or '<' or '~'
                || (value == '=' && options.Highlight)
                || (value == '+' && options.Inserted)
                || (value == '^' && options.Superscript)) {
                return false;
            }

            position++;
        }

        return foundToken || foundLiteralBracket;
    }

    private static bool TryParsePlainDoubleAsteriskInlines(
        string text,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        MarkdownInlineSourceMap? sourceMap,
        out InlineSequence sequence) {
        sequence = null!;
        if (state?.CaptureSyntaxTree != false
            || sourceMap != null
            || options.InlineParserExtensions.Count != 0
            || options.InlineTransformExtensions.Count != 0
            || (options.Abbreviations && state?.Abbreviations.Count > 0)
            || ContainsPotentialBareAutolinkSyntax(text, options)
            || options.CjkFriendlyEmphasis
            || string.IsNullOrEmpty(text)) {
            return false;
        }

        int runCount = 0;
        for (int position = 0; position < text.Length; position++) {
            char value = text[position];
            if (value == '*') {
                int runLength = 1;
                while (position + runLength < text.Length && text[position + runLength] == '*') {
                    runLength++;
                }

                if (runLength != 2) {
                    return false;
                }

                runCount++;
                position++;
                continue;
            }

            if (value is '\\' or '&' or '\r' or '\n' or '<' or '!' or '[' or ']' or '`' or '_' or '~' or '=' or '+' or '^') {
                return false;
            }
        }

        if (runCount < 2) {
            return false;
        }

        var runPositions = new int[runCount];
        var openerStack = new int[runCount];
        var closingRunByOpener = new int[runCount];
        for (int index = 0; index < closingRunByOpener.Length; index++) {
            closingRunByOpener[index] = -1;
        }

        int runIndex = 0;
        int openerCount = 0;
        int matchedCount = 0;
        for (int position = text.IndexOf('*'); position >= 0; position = text.IndexOf('*', position + 2)) {
            runPositions[runIndex] = position;
            GetDelimiterFlags(
                text,
                position,
                '*',
                2,
                options.CjkFriendlyEmphasis,
                out bool canOpen,
                out bool canClose);

            if (canClose && openerCount > 0) {
                int openerRun = openerStack[--openerCount];
                closingRunByOpener[openerRun] = runIndex;
                matchedCount++;
            } else if (canOpen) {
                openerStack[openerCount++] = runIndex;
            }

            runIndex++;
        }

        if (matchedCount == 0) {
            return false;
        }

        // Nested strong pairs need the full delimiter parser. This lightweight
        // path handles sequential pairs plus unmatched literal delimiter runs.
        int previousClosingPosition = -1;
        for (int openerRun = 0; openerRun < runCount; openerRun++) {
            int closingRun = closingRunByOpener[openerRun];
            if (closingRun < 0) {
                continue;
            }

            int openingPosition = runPositions[openerRun];
            if (openingPosition < previousClosingPosition) {
                return false;
            }

            previousClosingPosition = runPositions[closingRun] + 2;
        }

        sequence = new InlineSequence((matchedCount * 2) + 1) { AutoSpacing = false };
        int textStart = 0;
        for (int openerRun = 0; openerRun < runCount; openerRun++) {
            int closingRun = closingRunByOpener[openerRun];
            if (closingRun < 0) {
                continue;
            }

            int openingPosition = runPositions[openerRun];
            int closingPosition = runPositions[closingRun];
            AddSimpleTextRun(sequence, text, textStart, openingPosition - textStart);

            var strongContent = new InlineSequence(1) { AutoSpacing = false };
            AddSimpleTextRun(strongContent, text, openingPosition + 2, closingPosition - openingPosition - 2);
            sequence.AddRaw(new BoldSequenceInline(strongContent));
            textStart = closingPosition + 2;
        }

        AddSimpleTextRun(sequence, text, textStart, text.Length - textStart);
        return true;
    }

    private static bool IsPotentialInlineLinkOrImage(string text, int start) {
        int labelStart = text[start] == '!' ? start + 1 : start;
        if (labelStart >= text.Length || text[labelStart] != '[') {
            return false;
        }

        return text.IndexOf("](", labelStart + 1, StringComparison.Ordinal) >= 0;
    }

    private static bool TryReadSimpleCommonMarkInlineToken(
        string text,
        int start,
        MarkdownReaderOptions options,
        out SimpleInlineToken token) {
        token = default;
        if (start < 0 || start >= text.Length) {
            return false;
        }

        if (text[start] is '*' or '_') {
            char marker = text[start];
            int delimiterLength = start + 1 < text.Length && text[start + 1] == marker ? 2 : 1;
            if (start + delimiterLength + 1 >= text.Length) {
                return false;
            }

            int closing = delimiterLength == 2
                ? text.IndexOf(marker == '*' ? "**" : "__", start + delimiterLength, StringComparison.Ordinal)
                : text.IndexOf(marker, start + delimiterLength);
            if (closing <= start + delimiterLength
                || (delimiterLength == 1
                    && ((closing > 0 && text[closing - 1] == marker)
                        || (closing + 1 < text.Length && text[closing + 1] == marker)))
                || !IsSimpleEmphasisContent(text, start + delimiterLength, closing - start - delimiterLength, options)) {
                return false;
            }

            GetDelimiterFlags(text, start, marker, delimiterLength, options.CjkFriendlyEmphasis, out bool canOpen, out _);
            GetDelimiterFlags(text, closing, marker, delimiterLength, options.CjkFriendlyEmphasis, out _, out bool canClose);
            if (!canOpen || !canClose) {
                return false;
            }

            token = new SimpleInlineToken(
                delimiterLength == 2 ? SimpleInlineTokenKind.Strong : SimpleInlineTokenKind.Emphasis,
                start,
                closing + delimiterLength,
                start + delimiterLength,
                closing - start - delimiterLength);
            return true;
        }

        if (text[start] == '`') {
            if (start + 2 >= text.Length || text[start + 1] == '`') {
                return false;
            }

            int closing = text.IndexOf('`', start + 1);
            if (closing <= start + 1
                || char.IsWhiteSpace(text[start + 1])
                || char.IsWhiteSpace(text[closing - 1])
                || ContainsLineBreak(text, start + 1, closing - start - 1)) {
                return false;
            }

            token = new SimpleInlineToken(
                SimpleInlineTokenKind.Code,
                start,
                closing + 1,
                start + 1,
                closing - start - 1);
            return true;
        }

        bool isImage = text[start] == '!';
        int labelStart = isImage ? start + 1 : start;
        if (labelStart >= text.Length || text[labelStart] != '[') {
            return false;
        }

        int contentStart = labelStart + 1;
        int labelEnd = text.IndexOf(']', contentStart);
        if (labelEnd <= contentStart
            || labelEnd + 2 >= text.Length
            || text[labelEnd + 1] != '('
            || !IsSimpleInlineLiteral(text, contentStart, labelEnd - contentStart)) {
            return false;
        }

        int targetStart = labelEnd + 2;
        int targetEnd = text.IndexOf(')', targetStart);
        if (targetEnd < targetStart || !IsSimpleInlineLinkTarget(text, targetStart, targetEnd - targetStart)) {
            return false;
        }

        token = new SimpleInlineToken(
            isImage ? SimpleInlineTokenKind.Image : SimpleInlineTokenKind.Link,
            start,
            targetEnd + 1,
            contentStart,
            labelEnd - contentStart,
            targetStart,
            targetEnd - targetStart);
        return true;
    }

    private static bool ContainsLineBreak(string text, int start, int length) {
        int end = Math.Min(text.Length, start + length);
        for (int i = Math.Max(0, start); i < end; i++) {
            if (text[i] is '\r' or '\n') {
                return true;
            }
        }

        return false;
    }

    private static bool IsSimpleInlineLiteral(string text, int start, int length) {
        if (length <= 0 || start < 0 || start + length > text.Length) {
            return false;
        }

        int end = start + length;
        for (int i = start; i < end; i++) {
            char value = text[i];
            if (value is '\\' or '&' or '\r' or '\n' or '<' or '!' or '[' or '`' or '*' or '_' or '~' or '=' or '+' or '^') {
                return false;
            }
        }

        return true;
    }

    private static bool IsSimpleEmphasisContent(string text, int start, int length, MarkdownReaderOptions options) {
        if (length <= 0 || start < 0 || start + length > text.Length) {
            return false;
        }

        int end = start + length;
        for (int position = start; position < end; position++) {
            char value = text[position];
            if (value == '`') {
                if (!TryReadSimpleCommonMarkInlineToken(text, position, options, out var token)
                    || token.Kind != SimpleInlineTokenKind.Code
                    || token.End > end) {
                    return false;
                }

                position = token.End - 1;
                continue;
            }

            if (value is '\\' or '&' or '\r' or '\n' or '<' or '!' or '[' or '*' or '_' or '~' or '=' or '+' or '^') {
                return false;
            }
        }

        return true;
    }

    private static void AddSimpleEmphasisContent(
        InlineSequence sequence,
        string text,
        int start,
        int length,
        MarkdownReaderOptions options) {
        int end = start + length;
        int textStart = start;
        for (int position = start; position < end; position++) {
            if (text[position] != '`'
                || !TryReadSimpleCommonMarkInlineToken(text, position, options, out var token)
                || token.Kind != SimpleInlineTokenKind.Code
                || token.End > end) {
                continue;
            }

            AddSimpleTextRun(sequence, text, textStart, position - textStart);
            sequence.AddRaw(new CodeSpanInline(text.Substring(token.ContentStart, token.ContentLength)));
            position = token.End - 1;
            textStart = token.End;
        }

        AddSimpleTextRun(sequence, text, textStart, end - textStart);
    }

    private static bool IsSimpleInlineLinkTarget(string text, int start, int length) {
        if (length < 0 || start < 0 || start + length > text.Length) {
            return false;
        }

        int end = start + length;
        for (int i = start; i < end; i++) {
            char value = text[i];
            if (char.IsWhiteSpace(value)
                || value is '(' or ')' or '\\' or '<' or '>' or '"' or '\'' or '&') {
                return false;
            }
        }

        return true;
    }

    private static void AddSimpleTextRun(InlineSequence sequence, string text, int start, int length) {
        if (length > 0) {
            sequence.AddRaw(new MarkdownTextRun(text.Substring(start, length)));
        }
    }
}

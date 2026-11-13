namespace OfficeIMO.Markdown;

internal static partial class MarkdownReader {
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
        MarkdownLoadOptions options,
        MarkdownReaderState? state,
        MarkdownInlineSourceMap? sourceMap,
        out InlineSequence sequence) {
        sequence = null!;
        if (state?.CaptureSyntaxTree != false
            || sourceMap != null
            || options.InlineParserExtensions.Count != 0
            || options.Math
            || options.InlineTransformExtensions.Count != 0
            || (options.Abbreviations && state?.Abbreviations.Count > 0)
            || ContainsPotentialBareAutolinkSyntax(text, options)
            || string.IsNullOrEmpty(text)) {
            return false;
        }

        // Literal brackets are measured with the full parser's bracket lookups. A few lookups scan
        // the text directly; many of them stay linear only through the shared scan indexes.
        if (HasManyBracketOpeners(text)) {
            using var scanIndexScope = InlineScanIndexCache.Enter();
            return TryParseValidatedSimpleCommonMarkInlines(text, options, state, out sequence);
        }

        return TryParseValidatedSimpleCommonMarkInlines(text, options, state, out sequence);
    }

    private static bool HasManyBracketOpeners(string text) {
        const int DirectScanLimit = 8;
        int count = 0;
        for (int index = text.IndexOf('['); index >= 0; index = text.IndexOf('[', index + 1)) {
            if (++count > DirectScanLimit) {
                return true;
            }
        }

        return false;
    }

    private static bool TryParseValidatedSimpleCommonMarkInlines(
        string text,
        MarkdownLoadOptions options,
        MarkdownReaderState? state,
        out InlineSequence sequence) {
        sequence = null!;
        if (!ValidateSimpleCommonMarkInlineTokens(text, options, state)) {
            return false;
        }

        sequence = new InlineSequence { AutoSpacing = false };
        int position = 0;
        int textStart = 0;
        while (position < text.Length) {
            if (!TryReadSimpleCommonMarkInlineToken(text, position, options, out var token)) {
                int literalRunLength = GetSimpleLiteralRunLength(text, position, options, state);
                if (literalRunLength > 0) {
                    // Validation accepted this literal. The full parser writes it as its own text
                    // run, so the semantic document does too.
                    AddSimpleTextRun(sequence, text, textStart, position - textStart);
                    AddSimpleTextRun(sequence, text, position, literalRunLength);
                    position += literalRunLength;
                    textStart = position;
                    continue;
                }

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
        MarkdownLoadOptions options,
        MarkdownReaderState? state) {
        bool foundToken = false;
        int position = 0;
        while (position < text.Length) {
            char value = text[position];
            if (value is '*' or '_' or '`' or '[' or '!') {
                if (!TryReadSimpleCommonMarkInlineToken(text, position, options, out var token)) {
                    // A delimiter run that cannot open has no unresolved frame to affect on this
                    // validating path, and a bracket or '!' that forms no link or image is literal
                    // text. Anything else is left to the full parser.
                    if (value is '*' or '_') {
                        int runLength = GetDelimiterRunLength(text, position);
                        GetDelimiterFlags(
                            text,
                            position,
                            value,
                            runLength,
                            options.CjkFriendlyEmphasis,
                            out bool canOpen,
                            out _);
                        if (canOpen) {
                            return false;
                        }

                        foundToken = true;
                        position += runLength;
                        continue;
                    }

                    int literalLength = GetSimpleLiteralRunLength(text, position, options, state);
                    if (literalLength == 0) {
                        return false;
                    }

                    foundToken = true;
                    position += literalLength;
                    continue;
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

        return foundToken;
    }

    private static bool TryParsePlainDoubleAsteriskInlines(
        string text,
        MarkdownLoadOptions options,
        MarkdownReaderState? state,
        MarkdownInlineSourceMap? sourceMap,
        out InlineSequence sequence) {
        sequence = null!;
        if (state?.CaptureSyntaxTree != false
            || sourceMap != null
            || options.InlineParserExtensions.Count != 0
            || options.Math
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
            AddTextKeepingDoubleAsteriskRuns(sequence, text, textStart, openingPosition - textStart);

            var strongContent = new InlineSequence(1) { AutoSpacing = false };
            AddTextKeepingDoubleAsteriskRuns(strongContent, text, openingPosition + 2, closingPosition - openingPosition - 2);
            sequence.AddRaw(new BoldSequenceInline(strongContent));
            textStart = closingPosition + 2;
        }

        AddTextKeepingDoubleAsteriskRuns(sequence, text, textStart, text.Length - textStart);
        return true;
    }

    // Every "**" outside a matched pair is an unmatched run here, which the full parser keeps as
    // its own text run.
    private static void AddTextKeepingDoubleAsteriskRuns(InlineSequence sequence, string text, int start, int length) {
        int end = start + length;
        int textStart = start;
        for (int position = text.IndexOf('*', start); position >= 0 && position < end; position = text.IndexOf('*', position + 2)) {
            AddSimpleTextRun(sequence, text, textStart, position - textStart);
            AddSimpleTextRun(sequence, text, position, 2);
            textStart = position + 2;
        }

        AddSimpleTextRun(sequence, text, textStart, end - textStart);
    }

    private static bool TryReadSimpleCommonMarkInlineToken(
        string text,
        int start,
        MarkdownLoadOptions options,
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
                // A "**" pair matches only when both runs are exactly two long; longer runs take
                // part in the full parser's delimiter matching as a whole.
                || (delimiterLength == 2
                    && ((start + 2 < text.Length && text[start + 2] == marker)
                        || (closing + 2 < text.Length && text[closing + 2] == marker)))
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

    private static bool IsSimpleEmphasisContent(string text, int start, int length, MarkdownLoadOptions options) {
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
        MarkdownLoadOptions options) {
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

    // Length of the literal text run the full inline parser writes at this position, or 0 when it
    // would do something this path does not reproduce. Only called where no simple token starts.
    private static int GetSimpleLiteralRunLength(string text, int position, MarkdownLoadOptions options, MarkdownReaderState? state) {
        char value = text[position];
        switch (value) {
            case '*':
            case '_':
                return GetDelimiterRunLength(text, position);
            case '=' when !options.Highlight:
            case '+' when !options.Inserted:
            case '^' when !options.Superscript:
                return GetDelimiterRunLength(text, position);
            case '!':
                // "![" starts the image forms; a lone '!' is its own run.
                return position + 1 < text.Length && text[position + 1] == '[' ? 0 : 1;
            case '[':
                int bracketLength = GetSimpleLiteralBracketLength(text, position, options, state);
                // The literal is skipped whole, so it must not hide text that validation rejects.
                return bracketLength > 0 && ContainsSimpleRejectedCharacter(text, position, bracketLength, options) ? 0 : bracketLength;
            default:
                return 0;
        }
    }

    private static bool ContainsSimpleRejectedCharacter(string text, int start, int length, MarkdownLoadOptions options) {
        for (int i = start; i < start + length; i++) {
            char value = text[i];
            if (value is '\\' or '&' or '\n' or '<' or '~'
                || (value == '=' && options.Highlight)
                || (value == '+' && options.Inserted)
                || (value == '^' && options.Superscript)) {
                return true;
            }
        }

        return false;
    }

    private static int GetSimpleLiteralBracketLength(string text, int position, MarkdownLoadOptions options, MarkdownReaderState? state) {
        if (position + 1 < text.Length && (text[position + 1] == '^' || text[position + 1] == '!')) {
            // Footnote references and linked images.
            return 0;
        }

        var wrappers = InlineHtmlWrapperMatchIndex.Empty;
        if (TryParseLink(text, position, options, null, state, out _, out _, out _, out _, out _, out _, out _, out _, wrappers)) {
            return 0;
        }

        if (state == null) {
            return GetSimplePlainRunLength(text, position, options);
        }

        if (TryParseCollapsedRef(text, position, options, out _, out _, wrappers)
            || TryParseRefLink(text, position, options, out _, out _, out _, wrappers)) {
            return 0;
        }

        if (!TryParseShortcutRef(text, position, options, out int consumed, out var label, wrappers)) {
            return GetSimplePlainRunLength(text, position, options);
        }

        if (state.LinkRefs.ContainsKey(NormalizeReferenceLabel(label))
            || ContainsResolvedLinkInLabel(label, options, state)
            || (options.Abbreviations && ContainsAbbreviationCandidate(label, state))
            || ContainsBackslashEscapableCharacter(label)) {
            return 0;
        }

        // A bracketed label without a definition stays literal as one run.
        return consumed;
    }

    // A character no handler takes starts a text run that ends at the next possible inline start.
    private static int GetSimplePlainRunLength(string text, int position, MarkdownLoadOptions options) {
        int end = position + 1;
        while (end < text.Length && !IsPotentialInlineStart(text[end], options.InlineHtml, allowLinks: true, allowImages: true)) {
            end++;
        }

        return end - position;
    }

    private static int GetDelimiterRunLength(string text, int position) {
        char value = text[position];

        int runLength = 1;
        while (position + runLength < text.Length && text[position + runLength] == value) {
            runLength++;
        }

        return runLength;
    }

    private static void AddSimpleTextRun(InlineSequence sequence, string text, int start, int length) {
        if (length > 0) {
            sequence.AddRaw(new MarkdownTextRun(text.Substring(start, length)));
        }
    }
}

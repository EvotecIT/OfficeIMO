namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static bool TryConsumeStandaloneGenericAttributeBlock(
        string[] lines,
        int lineIndex,
        MarkdownReaderOptions options,
        MarkdownReaderState state) {
        if (!ShouldConsumeStandaloneGenericAttributeBlock(lines, lineIndex, options, state)
            || state.PendingGenericAttributeBlock != null
            || lines == null
            || lineIndex < 0
            || lineIndex >= lines.Length) {
            return false;
        }

        var line = lines[lineIndex] ?? string.Empty;
        var indentColumns = CountLeadingIndentColumns(line);
        if (indentColumns > 3) {
            return false;
        }

        var contentStart = CountLeadingSpaces(line);
        var content = line.Substring(contentStart).TrimEnd();
        if (!MarkdownGenericAttributeParser.TryConsumeLeadingAttributeBlock(
                content,
                out var remaining,
                out var attributes,
                out var consumedLength)
            || !string.IsNullOrWhiteSpace(remaining)
            || consumedLength != content.Length) {
            return false;
        }

        if (HasFollowingConsumedStandaloneAttributeTarget(lines, lineIndex, options)) {
            return true;
        }

        if (!HasFollowingSupportedStandaloneAttributeTarget(lines, lineIndex, options)) {
            return false;
        }

        var absoluteLine = state.SourceLineOffset + lineIndex + 1;
        var sourceSpan = CreateSpan(
            state,
            absoluteLine,
            contentStart + 1,
            absoluteLine,
            contentStart + consumedLength);

        state.PendingGenericAttributeBlock = new MarkdownPendingGenericAttributeBlock(
            attributes,
            content.Substring(0, consumedLength),
            sourceSpan);
        return true;
    }

    private static bool ShouldConsumeStandaloneGenericAttributeBlock(
        string[] lines,
        int lineIndex,
        MarkdownReaderOptions options,
        MarkdownReaderState state) =>
        ShouldParseBlockGenericAttributes(options, state)
        || IsQuoteStandaloneAttributeBeforeSupportedBlock(lines, lineIndex, options, state);

    private static bool IsQuoteStandaloneAttributeBeforeSupportedBlock(
        string[] lines,
        int lineIndex,
        MarkdownReaderOptions options,
        MarkdownReaderState state) {
        if (options?.GenericAttributes != true
            || state?.SuppressBlockGenericAttributes != true
            || lines == null
            || lineIndex < 0
            || lineIndex >= lines.Length
            || !state.QuoteContainerLines.Contains(lineIndex)) {
            return false;
        }

        for (int i = lineIndex + 1; i < lines.Length; i++) {
            if (string.IsNullOrWhiteSpace(lines[i])) {
                continue;
            }

            return state.QuoteContainerLines.Contains(i)
                && IsQuoteStandaloneAttributeSupportedBlock(lines, i, options, state);
        }

        return false;
    }

    private static bool IsQuoteStandaloneAttributeSupportedBlock(
        string[] lines,
        int lineIndex,
        MarkdownReaderOptions options,
        MarkdownReaderState state) {
        var line = lines[lineIndex];
        return (options.FencedCode && IsCodeFenceOpen(line, out _, out _, out _))
            || (options.UnorderedLists && IsUnorderedListLine(line, out _, out _, out _, out _))
            || (options.OrderedLists && IsOrderedListLine(line, out _, out _))
            || StartsTable(lines, lineIndex, options, state);
    }

    private static bool HasFollowingConsumedStandaloneAttributeTarget(
        string[] lines,
        int lineIndex,
        MarkdownReaderOptions options) {
        for (int i = lineIndex + 1; i < lines.Length; i++) {
            if (string.IsNullOrWhiteSpace(lines[i])) {
                continue;
            }

            return HtmlBlockParser.IsParagraphInterruptingHtmlBlockStart(lines[i], options)
                || IsStandaloneAttributeFootnoteDefinitionTarget(lines[i], options);
        }

        return false;
    }

    private static bool TryApplyPendingGenericAttributeBlock(
        MarkdownDoc document,
        int firstNewBlockIndex,
        int blockStartLine,
        MarkdownReaderState state,
        out int captureStartLine) {
        captureStartLine = -1;
        var pending = state.PendingGenericAttributeBlock;
        if (pending == null
            || document == null
            || firstNewBlockIndex < 0
            || firstNewBlockIndex >= document.Blocks.Count
            || document.Blocks[firstNewBlockIndex] is not MarkdownObject target) {
            return false;
        }

        if (target.Attributes.IsEmpty) {
            if (target is HeadingBlock heading) {
                heading.OffsetRelativeSourceInfoLines(Math.Max(0, blockStartLine + 1 - pending.SourceSpan.StartLine));
            }

            target.SetAttributes(pending.Attributes);
            MarkdownGenericAttributeSourceSpans.Set(target, pending.SourceText, pending.SourceSpan);
        }

        captureStartLine = pending.SourceSpan.StartLine - 1;
        state.PendingGenericAttributeBlock = null;
        return true;
    }

    private static bool TryTakePendingGenericAttributeBlock(
        MarkdownReaderState state,
        out MarkdownPendingGenericAttributeBlock pending) {
        pending = state.PendingGenericAttributeBlock!;
        if (pending == null) {
            return false;
        }

        state.PendingGenericAttributeBlock = null;
        return true;
    }

    private static bool HasFollowingSupportedStandaloneAttributeTarget(
        string[] lines,
        int lineIndex,
        MarkdownReaderOptions options) {
        for (int i = lineIndex + 1; i < lines.Length; i++) {
            if (string.IsNullOrWhiteSpace(lines[i])) {
                continue;
            }

            if (IsAtxHeading(lines[i], out _, out _)) {
                return true;
            }

            if (i + 1 < lines.Length
                && !string.IsNullOrWhiteSpace(lines[i + 1])
                && TryGetSetextHeadingUnderlineLevel(lines[i + 1], out _)) {
                return true;
            }

            if (options.FencedCode
                && IsCodeFenceOpen(lines[i], out _, out _, out _)) {
                return true;
            }

            if (options.UnorderedLists
                && IsUnorderedListLine(lines[i], out _, out _, out _, out _)) {
                return true;
            }

            if (options.OrderedLists
                && IsOrderedListLine(lines[i], out _, out _)) {
                return true;
            }

            if (StartsTable(lines, i, options)) {
                return true;
            }

            if (options.Headings
                && LooksLikeHr(lines[i])
                && TryGetSetextHeadingUnderlineLevel(lines[i], out _)) {
                return true;
            }

            if (options.IndentedCodeBlocks
                && CountLeadingIndentColumns(lines[i]) >= 4) {
                return true;
            }

            if (IsStandaloneAttributeReferenceDefinitionParagraphTarget(lines, i, options)) {
                return true;
            }

            if (options.Images
                && options.StandaloneImageBlocks
                && IsImageLine(lines[i])) {
                return true;
            }

            return options.Paragraphs
                && !IsCodeFenceOpen(lines[i], out _, out _, out _)
                && !StartsTable(lines, i, options)
                && !IsParagraphInterruptingThematicBreakLine(lines[i])
                && !IsParagraphInterruptingUnorderedListLine(lines[i])
                && !IsOrderedListLine(lines[i], out _, out _)
                && (!options.Callouts || !IsCalloutHeader(lines[i], options, out _, out _))
                && !IsQuoteStarter(lines[i])
                && !HtmlBlockParser.IsParagraphInterruptingHtmlBlockStart(lines[i], options)
                && !TryParseReferenceLinkDefinition(lines, i, options, out _, out _, out _, out _)
                && !(options.Abbreviations && TryParseAbbreviationDefinition(lines[i], 0, null, out _, out _, out _, out _, out _, out _))
                && !IsStandaloneAttributeFootnoteDefinitionTarget(lines[i], options)
                && !(options.StandaloneImageBlocks && IsImageLine(lines[i]));
        }

        return false;
    }

    private static bool IsStandaloneAttributeReferenceDefinitionParagraphTarget(
        string[] lines,
        int lineIndex,
        MarkdownReaderOptions options) =>
        options?.GenericAttributes == true
        && TryParseReferenceLinkDefinition(lines, lineIndex, options, out _, out _, out _, out _);

    private static bool IsStandaloneGenericAttributeOnlyLine(string? line) {
        if (string.IsNullOrWhiteSpace(line)) {
            return false;
        }

        var leading = CountLeadingSpaces(line!);
        var content = line!.Substring(leading).TrimEnd();
        return MarkdownGenericAttributeParser.TryConsumeLeadingAttributeBlock(
                content,
                out var remaining,
                out _,
                out var consumedLength)
            && consumedLength == content.Length
            && string.IsNullOrWhiteSpace(remaining);
    }

    private static bool IsReferenceDefinitionAfterStandaloneGenericAttribute(
        string[] lines,
        int lineIndex,
        MarkdownReaderOptions options) =>
        options?.GenericAttributes == true
        && lines != null
        && lineIndex > 0
        && lineIndex < lines.Length
        && IsStandaloneGenericAttributeOnlyLine(lines[lineIndex - 1])
        && TryParseReferenceLinkDefinition(lines, lineIndex, options, out _, out _, out _, out _);

    private static bool IsStandaloneAttributeFootnoteDefinitionTarget(string line, MarkdownReaderOptions options) {
        if (options?.Footnotes != true || string.IsNullOrWhiteSpace(line)) {
            return false;
        }

        var leading = 0;
        while (leading < line.Length && line[leading] == ' ') {
            leading++;
        }

        if (leading >= 4 || (leading < line.Length && line[leading] == '\t')) {
            return false;
        }

        var trimmed = line.TrimStart();
        if (trimmed.Length <= 4 || trimmed[0] != '[' || trimmed[1] != '^') {
            return false;
        }

        var closing = trimmed.IndexOf(']');
        return closing >= 2
            && closing + 1 < trimmed.Length
            && trimmed[closing + 1] == ':';
    }
}

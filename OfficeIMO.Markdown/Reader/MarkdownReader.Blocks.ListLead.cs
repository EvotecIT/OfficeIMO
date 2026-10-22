using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static ListItem CreateListItemFromLeadLines(List<string> lines, bool isTask, bool done, MarkdownReaderOptions options, MarkdownReaderState? state, int lineOffset, List<MarkdownSourceLineSlice>? sourceLines = null) {
        if (state?.IsListBoundaryProbe == true) {
            return isTask
                ? ListItem.TaskInlines(new InlineSequence(), done)
                : new ListItem(new InlineSequence());
        }

        if (state?.CaptureSyntaxTree == false
            && sourceLines == null
            && lines.Count == 1
            && !string.IsNullOrEmpty(lines[0])
            && !options.Abbreviations
            && !options.GenericAttributes
            && !StartsListItemLeadWithStandaloneBlock(lines, options)) {
            var content = ParseJoinedParagraphLines(lines, options, state);
            return isTask ? ListItem.TaskInlines(content, done) : new ListItem(content);
        }

        var leadingAbbreviationSyntax = RemoveLeadingListItemAbbreviationDefinitions(lines, options, state, lineOffset, sourceLines);

        ListItem AttachLeadingAbbreviationSyntax(ListItem item) {
            if (leadingAbbreviationSyntax.Count == 0) {
                return item;
            }

            item.SyntaxChildren.InsertRange(0, leadingAbbreviationSyntax);
            return item;
        }

        if (lines.Count == 0 || lines.TrueForAll(string.IsNullOrWhiteSpace)) {
            return AttachLeadingAbbreviationSyntax(isTask ? ListItem.TaskInlines(new InlineSequence(), done) : new ListItem(new InlineSequence()));
        }

        if (TryCreateListItemFromLeadBlocks(lines, isTask, done, options, state, sourceLines, out var blockLeadItem)) {
            return AttachLeadingAbbreviationSyntax(blockLeadItem);
        }

        if (TryParseListItemLeadSetextBlocks(lines, options, state, lineOffset, out var leadBlocks)) {
            var headingItem = isTask ? ListItem.TaskInlines(new InlineSequence(), done) : new ListItem(new InlineSequence());
            for (int i = 0; i < leadBlocks.Count; i++) {
                headingItem.NestedBlocks.Add(leadBlocks[i]);
            }
            return AttachLeadingAbbreviationSyntax(headingItem);
        }

        int firstBlank = lines.FindIndex(string.IsNullOrWhiteSpace);
        ParsedListItemGenericAttributes? listItemAttributes = TryConsumeListItemGenericAttributes(lines, sourceLines, options, state, lineOffset, firstBlank, out var parsedListItemAttributes)
            ? parsedListItemAttributes
            : null;

        ListItem AttachListItemAttributes(ListItem item) {
            if (item == null || listItemAttributes == null || listItemAttributes.Value.Attributes.IsEmpty) {
                return item!;
            }

            item.SetAttributes(listItemAttributes.Value.Attributes);
            item.GenericAttributeConsumedWhitespace = listItemAttributes.Value.ConsumedWhitespace;
            MarkdownGenericAttributeSourceSpans.Set(item, listItemAttributes.Value.SourceText, listItemAttributes.Value.SourceSpan);
            return item;
        }

        if (firstBlank <= 0) {
            var paragraphs = sourceLines != null && sourceLines.Count == lines.Count
                ? ParseParagraphsFromSourceLines(sourceLines, options, state)
                : ParseParagraphsFromLines(lines, options, state);
            var item = isTask ? ListItem.TaskInlines(paragraphs[0], done) : new ListItem(paragraphs[0]);
            for (int i = 1; i < paragraphs.Count; i++) {
                item.AdditionalParagraphs.Add(paragraphs[i]);
            }
            return AttachLeadingAbbreviationSyntax(AttachListItemAttributes(item));
        }

        var firstParagraph = sourceLines != null && sourceLines.Count >= firstBlank
            ? ParseParagraphsFromSourceLines(sourceLines.GetRange(0, firstBlank), options, state)[0]
            : ParseParagraphsFromLines(lines.GetRange(0, firstBlank), options, state)[0];
        var mixedItem = AttachListItemAttributes(isTask ? ListItem.TaskInlines(firstParagraph, done) : new ListItem(firstParagraph));

        if (firstBlank + 1 >= lines.Count) return AttachLeadingAbbreviationSyntax(mixedItem);

        var trailingLines = lines.GetRange(firstBlank + 1, lines.Count - firstBlank - 1);
        if (trailingLines.TrueForAll(string.IsNullOrWhiteSpace)) return AttachLeadingAbbreviationSyntax(mixedItem);

        if (sourceLines != null && sourceLines.Count > firstBlank + 1) {
            var trailingSourceLines = sourceLines.GetRange(firstBlank + 1, lines.Count - firstBlank - 1);
            if (!trailingSourceLines.TrueForAll(slice => string.IsNullOrWhiteSpace(slice.Text))) {
                var effectiveState = state ?? new MarkdownReaderState();
                var (trailingBlocksFromSource, trailingSyntaxFromSource) = ParseNestedMarkdownBlocks(trailingSourceLines, options, effectiveState);
                if (trailingSyntaxFromSource.All(node => node.Kind == MarkdownSyntaxKind.Paragraph)) {
                    var trailingParagraphs = ParseParagraphsFromSourceLines(trailingSourceLines, options, state);
                    for (int i = 0; i < trailingParagraphs.Count; i++) {
                        mixedItem.AdditionalParagraphs.Add(trailingParagraphs[i]);
                    }
                    return AttachLeadingAbbreviationSyntax(mixedItem);
                }

                for (int i = 0; i < trailingBlocksFromSource.Count; i++) {
                    mixedItem.NestedBlocks.Add(trailingBlocksFromSource[i]);
                }
                mixedItem.ForceLoose = true;
                return AttachLeadingAbbreviationSyntax(mixedItem);
            }
        }

        string[] trailingLineArray = trailingLines.ToArray();
        var trailingBlocks = ParseBlocksFromLines(trailingLineArray, options, state ?? new MarkdownReaderState());
        if (trailingBlocks.Count == 0
            && TryParseReferenceLinkDefinition(
                trailingLineArray,
                0,
                options,
                out _,
                out _,
                out _,
                out int consumedReferenceLines)
            && consumedReferenceLines == trailingLines.Count) {
            mixedItem.ForceLoose = true;
            return AttachLeadingAbbreviationSyntax(mixedItem);
        }

        if (mixedItem.TryAbsorbTrailingParagraphBlocks(trailingBlocks)) return AttachLeadingAbbreviationSyntax(mixedItem);

        for (int i = 0; i < trailingBlocks.Count; i++) {
            mixedItem.NestedBlocks.Add(trailingBlocks[i]);
        }
        mixedItem.ForceLoose = true;
        return AttachLeadingAbbreviationSyntax(mixedItem);
    }

    private static bool TryCreateSingleLineListItem(
        string line,
        bool isTask,
        bool done,
        MarkdownReaderOptions options,
        MarkdownReaderState state,
        int absoluteLine,
        int startColumn,
        out ListItem item) {
        item = null!;
        if (string.IsNullOrEmpty(line)
            || options.Abbreviations
            || options.GenericAttributes
            || StartsListItemLeadWithStandaloneBlock(line, options)) {
            return false;
        }

        string paragraph = GetParagraphLineJoinInfo(
            line,
            absoluteLine,
            startColumn,
            options,
            state.SourceTextMap,
            hasFollowingLine: false).Text;
        var sourceMap = state.CaptureSyntaxTree && state.SourceTextMap != null
            ? new MarkdownInlineSourceMap(state.SourceTextMap, paragraph, absoluteLine, startColumn)
            : null;
        var content = ParseInlines(paragraph, options, state, sourceMap);
        item = isTask ? ListItem.TaskInlines(content, done) : new ListItem(content);
        if (state.CaptureSyntaxTree) {
            item.SyntaxChildren.Add(BuildSyntaxNode(
                item.LeadParagraphBlock,
                CreateSpan(
                    state,
                    absoluteLine,
                    startColumn,
                    absoluteLine,
                    GetEndColumn(startColumn, paragraph))));
        }
        return true;
    }

    private readonly struct ParsedListItemGenericAttributes {
        public ParsedListItemGenericAttributes(MarkdownAttributeSet attributes, string sourceText, MarkdownSourceSpan? sourceSpan, string consumedWhitespace) {
            Attributes = attributes;
            SourceText = sourceText ?? string.Empty;
            SourceSpan = sourceSpan;
            ConsumedWhitespace = consumedWhitespace ?? string.Empty;
        }

        public MarkdownAttributeSet Attributes { get; }
        public string SourceText { get; }
        public MarkdownSourceSpan? SourceSpan { get; }
        public string ConsumedWhitespace { get; }
    }

    private static bool TryConsumeListItemGenericAttributes(
        List<string> lines,
        List<MarkdownSourceLineSlice>? sourceLines,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        int lineOffset,
        int firstBlank,
        out ParsedListItemGenericAttributes parsed) {
        parsed = default;
        if (options?.GenericAttributes != true || lines == null || lines.Count == 0) {
            return false;
        }

        int searchEnd = firstBlank > 0 ? firstBlank - 1 : lines.Count - 1;
        for (int lineIndex = searchEnd; lineIndex >= 0; lineIndex--) {
            var line = lines[lineIndex] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(line)) {
                continue;
            }

            if (!MarkdownGenericAttributeParser.TryConsumeTrailingAttributeBlock(
                    line,
                    out _,
                    out var attributes,
                    out var attributeStart,
                    out var attributeEnd,
                    requireLeadingWhitespace: true)) {
                return false;
            }

            var lineWithoutAttributeBlock = line.Substring(0, attributeStart);
            var trimmedTextLength = lineWithoutAttributeBlock.TrimEnd().Length;
            var consumedWhitespace = lineWithoutAttributeBlock.Substring(trimmedTextLength);
            var sourceText = line.Substring(attributeStart, attributeEnd - attributeStart + 1);
            var sourceLine = sourceLines != null && lineIndex >= 0 && lineIndex < sourceLines.Count
                ? sourceLines[lineIndex]
                : new MarkdownSourceLineSlice(line, (state?.SourceLineOffset ?? 0) + lineOffset + lineIndex + 1, 1);
            var attributeStartColumn = AdvanceSourceColumn(sourceLine.StartColumn, line, attributeStart);
            var attributeEndColumn = AdvanceSourceColumn(sourceLine.StartColumn, line, attributeEnd + 1) - 1;
            var span = CreateSpan(
                state,
                sourceLine.AbsoluteLine,
                attributeStartColumn,
                sourceLine.AbsoluteLine,
                attributeEndColumn);

            lines[lineIndex] = lineWithoutAttributeBlock;
            if (sourceLines != null && lineIndex >= 0 && lineIndex < sourceLines.Count) {
                sourceLines[lineIndex] = new MarkdownSourceLineSlice(
                    lineWithoutAttributeBlock,
                    sourceLine.AbsoluteLine,
                    sourceLine.StartColumn,
                    sourceLine.IsLazyQuoteContinuation,
                    sourceLine.IsQuoteContainerLine);
            }

            parsed = new ParsedListItemGenericAttributes(attributes, sourceText, span, consumedWhitespace);
            return true;
        }

        return false;
    }

    private static List<MarkdownSyntaxNode> RemoveLeadingListItemAbbreviationDefinitions(
        List<string> lines,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        int lineOffset,
        List<MarkdownSourceLineSlice>? sourceLines) {

        var syntaxNodes = new List<MarkdownSyntaxNode>();

        if (options?.Abbreviations != true || lines == null || lines.Count == 0) {
            return syntaxNodes;
        }

        int removeCount = 0;
        while (removeCount < lines.Count && IsAbbreviationDefinitionLine(lines[removeCount])) {
            if (TryBuildListItemAbbreviationDefinitionSyntaxNode(
                    lines[removeCount],
                    removeCount,
                    lineOffset,
                    state,
                    sourceLines,
                    out var syntaxNode)) {
                syntaxNodes.Add(syntaxNode);
            }

            removeCount++;
        }

        if (removeCount == 0) {
            return syntaxNodes;
        }

        lines.RemoveRange(0, removeCount);
        if (sourceLines != null && sourceLines.Count >= removeCount) {
            sourceLines.RemoveRange(0, removeCount);
        }

        return syntaxNodes;
    }

    private static bool TryBuildListItemAbbreviationDefinitionSyntaxNode(
        string line,
        int relativeLineIndex,
        int lineOffset,
        MarkdownReaderState? state,
        List<MarkdownSourceLineSlice>? sourceLines,
        out MarkdownSyntaxNode node) {
        node = null!;
        var effectiveState = state ?? new MarkdownReaderState();
        var sourceLine = sourceLines != null && relativeLineIndex >= 0 && relativeLineIndex < sourceLines.Count
            ? sourceLines[relativeLineIndex]
            : new MarkdownSourceLineSlice(line, (state?.SourceLineOffset ?? 0) + lineOffset + relativeLineIndex + 1, 1);
        var localLineIndex = sourceLine.AbsoluteLine - effectiveState.SourceLineOffset - 1;
        return TryBuildAbbreviationDefinitionSyntaxNode(
            sourceLine.Text,
            localLineIndex,
            sourceLine.StartColumn - 1,
            effectiveState,
            out node);
    }

    private static bool HasNonSyntaxOnlyListItemLeadSyntaxChildren(List<MarkdownSyntaxNode> syntaxChildren) {
        if (syntaxChildren == null || syntaxChildren.Count == 0) {
            return false;
        }

        for (int i = 0; i < syntaxChildren.Count; i++) {
            if (!IsListItemSyntaxOnlyDefinition(syntaxChildren[i])) {
                return true;
            }
        }

        return false;
    }

    private static bool IsListItemSyntaxOnlyDefinition(MarkdownSyntaxNode node) =>
        node != null
        && (node.Kind == MarkdownSyntaxKind.ReferenceLinkDefinition
            || node.Kind == MarkdownSyntaxKind.AbbreviationDefinition);

    private static bool TryCreateListItemFromLeadBlocks(
        List<string> lines,
        bool isTask,
        bool done,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        List<MarkdownSourceLineSlice>? sourceLines,
        out ListItem item) {
        item = null!;
        IReadOnlyList<IMarkdownBlock> leadBlocks;
        IReadOnlyList<MarkdownSyntaxNode> leadSyntax = Array.Empty<MarkdownSyntaxNode>();

        if (lines == null || lines.Count == 0 || !StartsListItemLeadWithStandaloneBlock(lines, options)) {
            return false;
        }

        if (sourceLines != null && sourceLines.Count == lines.Count) {
            var (blocksFromSource, syntaxFromSource) = ParseNestedMarkdownBlocks(sourceLines, options, state ?? new MarkdownReaderState());
            if (blocksFromSource.Count == 0 || blocksFromSource.All(block => block is ParagraphBlock)) {
                return false;
            }

            leadBlocks = blocksFromSource;
            leadSyntax = syntaxFromSource;
        } else if (!TryParseListItemLeadBlocks(lines, options, state, sourceLines, out leadBlocks)) {
            return false;
        }

        var blockLeadItem = isTask ? ListItem.TaskInlines(new InlineSequence(), done) : new ListItem(new InlineSequence());
        for (int i = 0; i < leadBlocks.Count; i++) {
            blockLeadItem.NestedBlocks.Add(leadBlocks[i]);
        }
        for (int i = 0; i < leadSyntax.Count; i++) {
            blockLeadItem.SyntaxChildren.Add(leadSyntax[i]);
        }
        if (leadBlocks.Count > 1 && lines.Exists(string.IsNullOrWhiteSpace)) {
            blockLeadItem.ForceLoose = true;
        }

        item = blockLeadItem;
        return true;
    }

    private static bool TryParseListItemLeadBlocks(
        List<string> lines,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        List<MarkdownSourceLineSlice>? sourceLines,
        out IReadOnlyList<IMarkdownBlock> blocks) {
        blocks = Array.Empty<IMarkdownBlock>();
        if (lines == null || lines.Count == 0) {
            return false;
        }

        if (!StartsListItemLeadWithStandaloneBlock(lines, options)) {
            return false;
        }

        IReadOnlyList<IMarkdownBlock> parsedBlocks;
        if (sourceLines != null && sourceLines.Count == lines.Count) {
            var (blocksFromSource, _) = ParseNestedMarkdownBlocks(sourceLines, options, state ?? new MarkdownReaderState());
            parsedBlocks = blocksFromSource;
        } else {
            parsedBlocks = ParseBlocksFromLines(lines.ToArray(), options, state ?? new MarkdownReaderState());
        }

        if (parsedBlocks.Count == 0 || parsedBlocks.All(block => block is ParagraphBlock)) {
            return false;
        }

        blocks = parsedBlocks;
        return true;
    }

    private static bool TryParseListItemLeadBlockSyntaxNodes(
        List<string> lines,
        int lineOffset,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        List<MarkdownSourceLineSlice>? sourceLines,
        out IReadOnlyList<MarkdownSyntaxNode> syntaxNodes) {
        syntaxNodes = Array.Empty<MarkdownSyntaxNode>();
        if (lines == null || lines.Count == 0) {
            return false;
        }

        if (!StartsListItemLeadWithStandaloneBlock(lines, options)) {
            return false;
        }

        IReadOnlyList<MarkdownSyntaxNode> parsedSyntax;
        if (sourceLines != null && sourceLines.Count == lines.Count) {
            parsedSyntax = ParseBlockSyntaxNodesFromSourceLines(sourceLines, options, state);
        } else {
            var leadSyntax = new List<MarkdownSyntaxNode>();
            ParseBlocksFromLines(lines.ToArray(), options, state ?? new MarkdownReaderState(), leadSyntax, lineOffset);
            parsedSyntax = leadSyntax;
        }

        if (parsedSyntax.Count == 0 || parsedSyntax.All(node => node.Kind == MarkdownSyntaxKind.Paragraph)) {
            return false;
        }

        syntaxNodes = parsedSyntax;
        return true;
    }

    private static bool StartsListItemLeadWithIndentedCode(List<string> lines, MarkdownReaderOptions options) {
        if (lines == null || lines.Count == 0 || options?.IndentedCodeBlocks != true) {
            return false;
        }

        int firstNonBlank = lines.FindIndex(line => !string.IsNullOrWhiteSpace(line));
        if (firstNonBlank < 0) {
            return false;
        }

        return CountLeadingIndentColumns(lines[firstNonBlank] ?? string.Empty) >= 4;
    }

    private static bool StartsListItemLeadWithStandaloneBlock(List<string> lines, MarkdownReaderOptions options) {
        if (lines == null || lines.Count == 0 || options == null) {
            return false;
        }

        if (StartsListItemLeadWithIndentedCode(lines, options)) {
            return true;
        }

        int firstNonBlank = lines.FindIndex(line => !string.IsNullOrWhiteSpace(line));
        if (firstNonBlank < 0) {
            return false;
        }

        return StartsListItemLeadWithStandaloneBlock(lines[firstNonBlank] ?? string.Empty, options);
    }

    private static bool StartsListItemLeadWithStandaloneBlock(string line, MarkdownReaderOptions options) {
        if (string.IsNullOrEmpty(line) || options == null) {
            return false;
        }

        if (options.IndentedCodeBlocks && CountLeadingIndentColumns(line) >= 4) {
            return true;
        }

        var trimmed = line.TrimStart();
        if (trimmed.Length == 0) {
            return false;
        }

        if (options.Headings && trimmed[0] == '#') {
            int markerLength = 0;
            while (markerLength < trimmed.Length && trimmed[markerLength] == '#') {
                markerLength++;
            }

            if (markerLength > 0 && markerLength <= 6 && (markerLength == trimmed.Length || char.IsWhiteSpace(trimmed[markerLength]))) {
                return true;
            }
        }

        if (trimmed[0] == '>') {
            return true;
        }

        if (options.FencedCode && IsCodeFenceOpen(trimmed, out _, out _, out _)) {
            return true;
        }

        if (options.UnorderedLists && IsUnorderedListLine(trimmed)) {
            return true;
        }

        if (options.OrderedLists && IsOrderedListLine(trimmed, options, out _, out _)) {
            return true;
        }

        return false;
    }

    private static bool TryParseListItemLeadSetextBlocks(List<string> lines, MarkdownReaderOptions options, MarkdownReaderState? state, int lineOffset, out List<IMarkdownBlock> blocks) {
        blocks = new List<IMarkdownBlock>();
        if (lines == null || lines.Count == 0 || options == null || !options.Headings) return false;

        if (!TryGetLeadingSetextHeadingPrefix(lines, options, out int headingLineCount, out int level, out string headingText, out var headingSourceMap)) return false;
        if (IsSetextHeadingUnderlineSuppressed(state, lineOffset + headingLineCount - 1)) return false;

        blocks.Add(new HeadingBlock(level, ParseInlines(headingText, options, state, headingSourceMap)));

        if (headingLineCount >= lines.Count) return true;

        var trailingLines = lines.GetRange(headingLineCount, lines.Count - headingLineCount);
        if (trailingLines.TrueForAll(string.IsNullOrWhiteSpace)) return true;

        var trailingBlocks = ParseBlocksFromLines(trailingLines.ToArray(), options, state ?? new MarkdownReaderState());
        for (int i = 0; i < trailingBlocks.Count; i++) {
            blocks.Add(trailingBlocks[i]);
        }

        return true;
    }

    private static bool TryGetLeadingSetextHeadingPrefix(List<string> lines, MarkdownReaderOptions options, out int headingLineCount, out int level, out string headingText, out MarkdownInlineSourceMap? headingSourceMap) {
        headingLineCount = 0;
        level = 0;
        headingText = string.Empty;
        headingSourceMap = null;
        if (lines == null || lines.Count < 2 || options == null || !options.Headings) return false;

        int firstBlank = lines.FindIndex(string.IsNullOrWhiteSpace);
        int maxPrefixLength = firstBlank >= 0 ? firstBlank : lines.Count;
        if (maxPrefixLength < 2) return false;

        for (int underlineIndex = 1; underlineIndex < maxPrefixLength; underlineIndex++) {
            if (!TryGetSetextHeadingUnderlineLevel(lines[underlineIndex] ?? string.Empty, out level)) {
                continue;
            }

            var contentLines = lines.GetRange(0, underlineIndex);
            if (contentLines.TrueForAll(string.IsNullOrWhiteSpace)) {
                level = 0;
                return false;
            }

            (headingText, headingSourceMap) = JoinTrimmedParagraphLines(contentLines, options);
            if (headingText.Length == 0) {
                level = 0;
                return false;
            }

            headingLineCount = underlineIndex + 1;
            return true;
        }

        level = 0;
        headingText = string.Empty;
        return false;
    }

    private static bool TryParseSetextHeadingParagraphLines(List<string> lines, MarkdownReaderOptions options, out int level, out string headingText) {
        level = 0;
        headingText = string.Empty;
        if (lines == null || lines.Count < 2 || options == null || !options.Headings) return false;

        var underlineLine = lines[lines.Count - 1] ?? string.Empty;
        if (!TryGetSetextHeadingUnderlineLevel(underlineLine, out level)) return false;

        var contentLines = lines.GetRange(0, lines.Count - 1);
        if (contentLines.Count == 0 || contentLines.TrueForAll(string.IsNullOrWhiteSpace)) return false;

        headingText = JoinTrimmedParagraphLines(contentLines, options).Text;
        return headingText.Length > 0;
    }

}

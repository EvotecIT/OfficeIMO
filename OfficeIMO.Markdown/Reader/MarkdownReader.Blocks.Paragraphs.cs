using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static List<IMarkdownBlock> ParseBlocksFromLines(string[] lines, MarkdownReaderOptions options, MarkdownReaderState state, List<MarkdownSyntaxNode>? syntaxNodes = null, int lineOffset = 0) {
        var doc = MarkdownDoc.Create();
        var opt = CloneOptionsWithoutFrontMatter(options);
        var pipeline = MarkdownReaderPipeline.Default(opt);
        int previousLineOffset = state.SourceLineOffset;
        state.SourceLineOffset = lineOffset;

        try {
            int i = 0;
            while (i < lines.Length) {
                if (string.IsNullOrWhiteSpace(lines[i])) { i++; continue; }
                bool matched = false;
                var parsers = pipeline.Parsers;
                int previousBlockCount = doc.Blocks.Count;
                int startLine = lineOffset + i;
                for (int p = 0; p < parsers.Count; p++) {
                    if (parsers[p].TryParse(lines, ref i, opt, doc, state)) {
                        matched = true;
                        if (syntaxNodes != null && doc.Blocks.Count > previousBlockCount) {
                            CaptureSyntaxNodes(doc, previousBlockCount, startLine, lineOffset + i, syntaxNodes, state);
                        }
                        break;
                    }
                }
                if (!matched) i++;
            }
        } finally {
            state.SourceLineOffset = previousLineOffset;
        }

        return doc.Blocks.ToList();
    }

    private static bool EndsWithTwoSpacesLine(string s) {
        if (string.IsNullOrEmpty(s)) return false;
        int n = s.Length - 1;
        int count = 0;
        while (n >= 0 && s[n] == ' ') {
            count++;
            n--;
            if (count >= 2) return true;
        }
        return false;
    }

    private readonly struct MarkdownSourceLineSlice {
        public MarkdownSourceLineSlice(string text, int absoluteLine, int startColumn, bool isLazyQuoteContinuation = false, bool isQuoteContainerLine = false) {
            Text = text ?? string.Empty;
            AbsoluteLine = absoluteLine;
            StartColumn = startColumn < 1 ? 1 : startColumn;
            IsLazyQuoteContinuation = isLazyQuoteContinuation;
            IsQuoteContainerLine = isQuoteContainerLine;
        }

        public string Text { get; }
        public int AbsoluteLine { get; }
        public int StartColumn { get; }
        public bool IsLazyQuoteContinuation { get; }
        public bool IsQuoteContainerLine { get; }
    }

    private readonly struct ParagraphLineJoinInfo {
        public ParagraphLineJoinInfo(
            string text,
            bool hardBreak,
            string? hardBreakMarker,
            MarkdownSourceSpan? hardBreakMarkerSpan,
            bool preserveLineBreak = false) {
            Text = text ?? string.Empty;
            HardBreak = hardBreak;
            HardBreakMarker = hardBreakMarker;
            HardBreakMarkerSpan = hardBreakMarkerSpan;
            PreserveLineBreak = preserveLineBreak;
        }

        public string Text { get; }
        public bool HardBreak { get; }
        public string? HardBreakMarker { get; }
        public MarkdownSourceSpan? HardBreakMarkerSpan { get; }
        public bool PreserveLineBreak { get; }
        public bool UsesLineBreakSeparator => HardBreak || PreserveLineBreak;
    }

    private static List<InlineSequence> ParseParagraphsFromLines(List<string> lines, MarkdownReaderOptions options, MarkdownReaderState? state) {
        var paragraphs = new List<InlineSequence>();
        if (lines == null || lines.Count == 0) {
            paragraphs.Add(ParseInlines(string.Empty, options, state));
            return paragraphs;
        }

        var cur = new List<string>();
        for (int i = 0; i < lines.Count; i++) {
            var ln = lines[i] ?? string.Empty;
            if (ln.Length == 0) {
                if (cur.Count > 0) {
                    paragraphs.Add(ParseInlines(JoinParagraphLines(cur, options), options, state));
                    cur.Clear();
                }
                continue;
            }
            cur.Add(ln);
        }
        if (cur.Count > 0) paragraphs.Add(ParseInlines(JoinParagraphLines(cur, options), options, state));

        if (paragraphs.Count == 0) paragraphs.Add(ParseInlines(string.Empty, options, state));
        return paragraphs;
    }

    private static List<InlineSequence> ParseParagraphsFromSourceLines(List<MarkdownSourceLineSlice> lines, MarkdownReaderOptions options, MarkdownReaderState? state) {
        var paragraphs = new List<InlineSequence>();
        if (lines == null || lines.Count == 0) {
            paragraphs.Add(ParseInlines(string.Empty, options, state));
            return paragraphs;
        }

        var current = new List<MarkdownSourceLineSlice>();
        for (int i = 0; i < lines.Count; i++) {
            if (string.IsNullOrEmpty(lines[i].Text)) {
                if (current.Count > 0) {
                    var (text, sourceMap) = JoinParagraphSourceLinesWithSourceMap(current, options, state);
                    paragraphs.Add(ParseInlines(text, options, state, sourceMap));
                    current.Clear();
                }
                continue;
            }

            current.Add(lines[i]);
        }

        if (current.Count > 0) {
            var (text, sourceMap) = JoinParagraphSourceLinesWithSourceMap(current, options, state);
            paragraphs.Add(ParseInlines(text, options, state, sourceMap));
        }

        if (paragraphs.Count == 0) {
            paragraphs.Add(ParseInlines(string.Empty, options, state));
        }

        return paragraphs;
    }

    private static List<ParagraphBlock> ParseParagraphBlocksFromLines(List<string> lines, MarkdownReaderOptions options, MarkdownReaderState? state) {
        var paragraphInlines = ParseParagraphsFromLines(lines, options, state);
        var blocks = new List<ParagraphBlock>(paragraphInlines.Count);
        for (int i = 0; i < paragraphInlines.Count; i++) {
            blocks.Add(new ParagraphBlock(paragraphInlines[i]));
        }
        return blocks;
    }

    private static List<ParagraphBlock> ParseParagraphBlocksFromSourceLines(List<MarkdownSourceLineSlice> lines, MarkdownReaderOptions options, MarkdownReaderState? state) {
        var paragraphInlines = ParseParagraphsFromSourceLines(lines, options, state);
        var blocks = new List<ParagraphBlock>(paragraphInlines.Count);
        for (int i = 0; i < paragraphInlines.Count; i++) {
            blocks.Add(new ParagraphBlock(paragraphInlines[i]));
        }
        return blocks;
    }

    private static void AddListItemLeadSyntaxNodes(ListItem item, List<string> lines, int lineOffset, MarkdownReaderOptions options, MarkdownReaderState? state, List<MarkdownSourceLineSlice>? sourceLines = null) {
        if (item == null || lines == null || lines.Count == 0) return;
        if (HasNonSyntaxOnlyListItemLeadSyntaxChildren(item.SyntaxChildren)) return;
        int absoluteLineOffset = (state?.SourceLineOffset ?? 0) + lineOffset;

        if (TryParseListItemLeadBlockSyntaxNodes(lines, lineOffset, options, state, sourceLines, out var leadBlockSyntax)) {
            for (int i = 0; i < leadBlockSyntax.Count; i++) {
                item.SyntaxChildren.Add(leadBlockSyntax[i]);
            }
            return;
        }

        if (TryParseSetextHeadingParagraphLines(lines, options, out int level, out string headingText)) {
            item.SyntaxChildren.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.Heading,
                CreateLineSpan(state, absoluteLineOffset + 1, absoluteLineOffset + lines.Count),
                headingText));
            return;
        }

        if (TryGetLeadingSetextHeadingPrefix(lines, options, out int headingLineCount, out level, out headingText)) {
            item.SyntaxChildren.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.Heading,
                CreateLineSpan(state, absoluteLineOffset + 1, absoluteLineOffset + headingLineCount),
                headingText));

            if (headingLineCount < lines.Count) {
                var trailingLines = lines.GetRange(headingLineCount, lines.Count - headingLineCount);
                if (!trailingLines.TrueForAll(string.IsNullOrWhiteSpace)) {
                    IReadOnlyList<MarkdownSyntaxNode> trailingSyntax;
                    if (sourceLines != null && sourceLines.Count >= headingLineCount) {
                        trailingSyntax = ParseBlockSyntaxNodesFromSourceLines(sourceLines.GetRange(headingLineCount, lines.Count - headingLineCount), options, state);
                    } else {
                        var nestedSyntax = new List<MarkdownSyntaxNode>();
                        ParseBlocksFromLines(trailingLines.ToArray(), options, state ?? new MarkdownReaderState(), nestedSyntax, lineOffset: lineOffset + headingLineCount);
                        trailingSyntax = nestedSyntax;
                    }

                    for (int i = 0; i < trailingSyntax.Count; i++) {
                        item.SyntaxChildren.Add(trailingSyntax[i]);
                    }
                }
            }
            return;
        }

        int firstBlank = lines.FindIndex(string.IsNullOrWhiteSpace);
        if (firstBlank > 0) {
            if (sourceLines != null && sourceLines.Count >= firstBlank) {
                AddParagraphSyntaxNodes(item.SyntaxChildren, sourceLines.GetRange(0, firstBlank), options, state);
            } else {
                AddParagraphSyntaxNodes(item.SyntaxChildren, lines.GetRange(0, firstBlank), absoluteLineOffset, options, state);
            }

            if (firstBlank + 1 < lines.Count) {
                var trailingLines = lines.GetRange(firstBlank + 1, lines.Count - firstBlank - 1);
                if (!trailingLines.TrueForAll(string.IsNullOrWhiteSpace)) {
                    IReadOnlyList<MarkdownSyntaxNode> trailingSyntax;
                    if (sourceLines != null && sourceLines.Count > firstBlank + 1) {
                        trailingSyntax = ParseBlockSyntaxNodesFromSourceLines(sourceLines.GetRange(firstBlank + 1, lines.Count - firstBlank - 1), options, state);
                    } else {
                        var nestedSyntax = new List<MarkdownSyntaxNode>();
                        ParseBlocksFromLines(trailingLines.ToArray(), options, state ?? new MarkdownReaderState(), nestedSyntax, lineOffset: lineOffset + firstBlank + 1);
                        trailingSyntax = nestedSyntax;
                    }

                    for (int i = 0; i < trailingSyntax.Count; i++) {
                        item.SyntaxChildren.Add(trailingSyntax[i]);
                    }
                    return;
                }
            }

            return;
        }

        if (sourceLines != null && sourceLines.Count == lines.Count) {
            AddParagraphSyntaxNodes(item.SyntaxChildren, sourceLines, options, state);
        } else {
            AddParagraphSyntaxNodes(item.SyntaxChildren, lines, absoluteLineOffset, options, state);
        }
    }

    private static void AddParagraphSyntaxNodes(List<MarkdownSyntaxNode> nodes, List<string> lines, int lineOffset, MarkdownReaderOptions options, MarkdownReaderState? state) {
        if (nodes == null || lines == null || lines.Count == 0) return;

        var current = new List<string>();
        int currentStart = -1;

        for (int i = 0; i < lines.Count; i++) {
            var line = lines[i] ?? string.Empty;
            if (line.Length == 0) {
                FlushParagraphSyntaxNode(nodes, current, currentStart, i - 1, lineOffset, options, state);
                current.Clear();
                currentStart = -1;
                continue;
            }

            if (currentStart < 0) currentStart = i;
            current.Add(line);
        }

        FlushParagraphSyntaxNode(nodes, current, currentStart, lines.Count - 1, lineOffset, options, state);
    }

    private static void AddParagraphSyntaxNodes(List<MarkdownSyntaxNode> nodes, List<MarkdownSourceLineSlice> lines, MarkdownReaderOptions options, MarkdownReaderState? state) {
        if (nodes == null || lines == null || lines.Count == 0) return;

        var current = new List<MarkdownSourceLineSlice>();
        for (int i = 0; i < lines.Count; i++) {
            if (string.IsNullOrEmpty(lines[i].Text)) {
                FlushParagraphSyntaxNode(nodes, current, options, state);
                current.Clear();
                continue;
            }

            current.Add(lines[i]);
        }

        FlushParagraphSyntaxNode(nodes, current, options, state);
    }

    private static IReadOnlyList<MarkdownSyntaxNode> ParseBlockSyntaxNodesFromSourceLines(
        List<MarkdownSourceLineSlice> lines,
        MarkdownReaderOptions options,
        MarkdownReaderState? state) {
        if (lines == null || lines.Count == 0) {
            return Array.Empty<MarkdownSyntaxNode>();
        }

        var effectiveState = state ?? new MarkdownReaderState();
        var (_, syntaxChildren) = ParseNestedMarkdownBlocks(lines, options, effectiveState);
        return syntaxChildren;
    }

    private static void FlushParagraphSyntaxNode(List<MarkdownSyntaxNode> nodes, List<string> lines, int startIndex, int endIndex, int lineOffset, MarkdownReaderOptions options, MarkdownReaderState? state) {
        if (nodes == null || lines == null || lines.Count == 0 || startIndex < 0 || endIndex < startIndex) return;

        var inlines = ParseInlines(JoinParagraphLines(lines, options), options, state);
        var paragraph = new ParagraphBlock(inlines);
        nodes.Add(BuildSyntaxNode(paragraph, CreateLineSpan(state, lineOffset + startIndex + 1, lineOffset + endIndex + 1)));
    }

    private static void FlushParagraphSyntaxNode(List<MarkdownSyntaxNode> nodes, List<MarkdownSourceLineSlice> lines, MarkdownReaderOptions options, MarkdownReaderState? state) {
        if (nodes == null || lines == null || lines.Count == 0) return;

        var (text, sourceMap) = JoinParagraphSourceLinesWithSourceMap(lines, options, state);
        var inlines = ParseInlines(text, options, state, sourceMap);
        var paragraph = new ParagraphBlock(inlines);
        nodes.Add(BuildSyntaxNode(paragraph, CreateSpan(
            state,
            lines[0].AbsoluteLine,
            lines[0].StartColumn,
            lines[lines.Count - 1].AbsoluteLine,
            lines[lines.Count - 1].StartColumn + Math.Max(0, lines[lines.Count - 1].Text.Length - 1))));
    }

    private static void AddListItemChildSyntaxNode(ListItem item, IMarkdownBlock block, int startLineIndex, int endExclusiveLineIndex, MarkdownReaderState? state) {
        if (item == null || block == null) return;
        int absoluteStart = (state?.SourceLineOffset ?? 0) + startLineIndex;
        int absoluteEndExclusive = (state?.SourceLineOffset ?? 0) + endExclusiveLineIndex;
        item.SyntaxChildren.Add(BuildSyntaxNode(block, CreateLineSpan(state, absoluteStart + 1, Math.Max(absoluteStart + 1, absoluteEndExclusive))));
    }

    private static void AddListItemChildSyntaxNode(
        ListItem item,
        IMarkdownBlock block,
        string[] sourceLines,
        int continuationIndent,
        int startLineIndex,
        int endExclusiveLineIndex,
        MarkdownReaderState? state) {

        if (item == null || block == null || sourceLines == null) return;

        var slices = BuildListItemNestedSourceLines(sourceLines, continuationIndent, startLineIndex, endExclusiveLineIndex, state);
        if (slices.Count == 0) {
            AddListItemChildSyntaxNode(item, block, startLineIndex, endExclusiveLineIndex, state);
            return;
        }

        if (block is QuoteBlock quoteBlock && quoteBlock.SourceSpan.HasValue) {
            var mappedNode = BuildSyntaxNode(quoteBlock, quoteBlock.SourceSpan);
            SynchronizeOwnedSyntaxCaches(mappedNode);
            MarkdownObjectTreeBinder.BindSourceSpans(mappedNode);
            item.SyntaxChildren.Add(mappedNode);
            return;
        }

        if (block is CustomContainerBlock customContainerBlock && customContainerBlock.OpeningFenceSourceSpan.HasValue) {
            var opening = customContainerBlock.OpeningFenceSourceSpan.Value;
            var closing = customContainerBlock.ClosingFenceSourceSpan ?? customContainerBlock.SourceSpan ?? opening;
            var sourceSpan = new MarkdownSourceSpan(
                opening.StartLine,
                opening.StartColumn ?? 1,
                closing.EndLine,
                closing.EndColumn ?? opening.EndColumn ?? 1);
            var mappedNode = BuildSyntaxNode(customContainerBlock, sourceSpan);
            SynchronizeOwnedSyntaxCaches(mappedNode);
            MarkdownObjectTreeBinder.BindSourceSpans(mappedNode);
            item.SyntaxChildren.Add(mappedNode);
            return;
        }

        var lastLine = slices[slices.Count - 1].Text;
        var localSpan = new MarkdownSourceSpan(1, 1, slices.Count, Math.Max(1, lastLine.Length));
        var markdownObject = block as MarkdownObject;
        var originalAttributeSourceText = MarkdownGenericAttributeSourceSpans.GetSourceText(markdownObject);
        var originalAttributeSourceSpan = MarkdownGenericAttributeSourceSpans.GetSourceSpan(markdownObject);
        MarkdownSourceSpan localAttributeSourceSpan = default;
        var useLocalAttributeSourceSpan = markdownObject != null
            && originalAttributeSourceSpan.HasValue
            && TryMapSourceSpanToNestedLocal(slices, originalAttributeSourceSpan.Value, out localAttributeSourceSpan);
        if (useLocalAttributeSourceSpan) {
            MarkdownGenericAttributeSourceSpans.Set(markdownObject, originalAttributeSourceText, localAttributeSourceSpan);
        }

        MarkdownSyntaxNode localNode;
        try {
            localNode = BuildSyntaxNode(block, localSpan);
        } finally {
            if (useLocalAttributeSourceSpan) {
                MarkdownGenericAttributeSourceSpans.Set(markdownObject, originalAttributeSourceText, originalAttributeSourceSpan);
            }
        }

        var remappedNode = RemapNestedSyntaxNode(slices, localNode);
        SynchronizeOwnedSyntaxCaches(remappedNode);
        MarkdownObjectTreeBinder.BindSourceSpans(remappedNode);
        item.SyntaxChildren.Add(remappedNode);
    }

    private static bool TryMapSourceSpanToNestedLocal(
        IReadOnlyList<MarkdownSourceLineSlice> sourceLines,
        MarkdownSourceSpan sourceSpan,
        out MarkdownSourceSpan localSpan) {
        localSpan = default;
        if (sourceLines == null || sourceLines.Count == 0 || !sourceSpan.StartColumn.HasValue) {
            return false;
        }

        var startIndex = FindNestedSourceLineIndex(sourceLines, sourceSpan.StartLine);
        var endIndex = FindNestedSourceLineIndex(sourceLines, sourceSpan.EndLine);
        if (startIndex < 0 || endIndex < 0) {
            return false;
        }

        var startLine = sourceLines[startIndex];
        var endLine = sourceLines[endIndex];
        var startColumn = Math.Max(1, sourceSpan.StartColumn.Value - startLine.StartColumn + 1);
        var endColumn = sourceSpan.EndColumn.HasValue
            ? Math.Max(1, sourceSpan.EndColumn.Value - endLine.StartColumn + 1)
            : Math.Max(1, endLine.Text.Length);

        localSpan = new MarkdownSourceSpan(startIndex + 1, startColumn, endIndex + 1, endColumn);
        return true;
    }

    private static int FindNestedSourceLineIndex(IReadOnlyList<MarkdownSourceLineSlice> sourceLines, int absoluteLine) {
        for (var i = 0; i < sourceLines.Count; i++) {
            if (sourceLines[i].AbsoluteLine == absoluteLine) {
                return i;
            }
        }

        return -1;
    }

    private static List<MarkdownSourceLineSlice> BuildListItemNestedSourceLines(
        string[] sourceLines,
        int continuationIndent,
        int startLineIndex,
        int endExclusiveLineIndex,
        MarkdownReaderState? state) {

        var count = Math.Max(0, Math.Min(endExclusiveLineIndex, sourceLines.Length) - Math.Max(0, startLineIndex));
        var slices = new List<MarkdownSourceLineSlice>(count);
        var start = Math.Max(0, startLineIndex);
        var end = Math.Min(endExclusiveLineIndex, sourceLines.Length);

        for (var i = start; i < end; i++) {
            var line = sourceLines[i] ?? string.Empty;
            var absoluteLine = GetSourceLineAbsoluteNumber(state, i);
            if (string.IsNullOrWhiteSpace(line)) {
                slices.Add(new MarkdownSourceLineSlice(string.Empty, absoluteLine, 1));
                continue;
            }

            if (CountLeadingIndentColumns(line) >= continuationIndent) {
                slices.Add(new MarkdownSourceLineSlice(
                    StripLeadingIndentColumns(line, continuationIndent),
                    absoluteLine,
                    continuationIndent + 1));
                continue;
            }

            var leadingColumns = CountLeadingIndentColumns(line);
            slices.Add(new MarkdownSourceLineSlice(
                line.TrimStart(),
                absoluteLine,
                leadingColumns + 1));
        }

        return slices;
    }

    private static int GetSourceLineAbsoluteNumber(MarkdownReaderState? state, int lineIndex) {
        if (state?.SourceLineAbsoluteNumbers != null &&
            lineIndex >= 0 &&
            lineIndex < state.SourceLineAbsoluteNumbers.Count) {
            return state.SourceLineAbsoluteNumbers[lineIndex];
        }

        return (state?.SourceLineOffset ?? 0) + lineIndex + 1;
    }

    private static ListItem CreateListItemFromLeadLines(List<string> lines, bool isTask, bool done, MarkdownReaderOptions options, MarkdownReaderState? state, int lineOffset, List<MarkdownSourceLineSlice>? sourceLines = null) {
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

        var trailingBlocks = ParseBlocksFromLines(trailingLines.ToArray(), options, state ?? new MarkdownReaderState());
        if (mixedItem.TryAbsorbTrailingParagraphBlocks(trailingBlocks)) return AttachLeadingAbbreviationSyntax(mixedItem);

        for (int i = 0; i < trailingBlocks.Count; i++) {
            mixedItem.NestedBlocks.Add(trailingBlocks[i]);
        }
        mixedItem.ForceLoose = true;
        return AttachLeadingAbbreviationSyntax(mixedItem);
    }

    private readonly struct ParsedListItemGenericAttributes {
        public ParsedListItemGenericAttributes(MarkdownAttributeSet attributes, string sourceText, MarkdownSourceSpan sourceSpan, string consumedWhitespace) {
            Attributes = attributes;
            SourceText = sourceText ?? string.Empty;
            SourceSpan = sourceSpan;
            ConsumedWhitespace = consumedWhitespace ?? string.Empty;
        }

        public MarkdownAttributeSet Attributes { get; }
        public string SourceText { get; }
        public MarkdownSourceSpan SourceSpan { get; }
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

        var firstLine = lines[firstNonBlank] ?? string.Empty;
        var trimmed = firstLine.TrimStart();
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

        if (options.UnorderedLists && IsUnorderedListLine(trimmed, out _, out _, out _)) {
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

        if (!TryGetLeadingSetextHeadingPrefix(lines, options, out int headingLineCount, out int level, out string headingText)) return false;
        if (IsSetextHeadingUnderlineSuppressed(state, lineOffset + headingLineCount - 1)) return false;

        blocks.Add(new HeadingBlock(level, ParseInlines(headingText, options, state)));

        if (headingLineCount >= lines.Count) return true;

        var trailingLines = lines.GetRange(headingLineCount, lines.Count - headingLineCount);
        if (trailingLines.TrueForAll(string.IsNullOrWhiteSpace)) return true;

        var trailingBlocks = ParseBlocksFromLines(trailingLines.ToArray(), options, state ?? new MarkdownReaderState());
        for (int i = 0; i < trailingBlocks.Count; i++) {
            blocks.Add(trailingBlocks[i]);
        }

        return true;
    }

    private static bool TryGetLeadingSetextHeadingPrefix(List<string> lines, MarkdownReaderOptions options, out int headingLineCount, out int level, out string headingText) {
        headingLineCount = 0;
        level = 0;
        headingText = string.Empty;
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

            headingText = JoinParagraphLines(contentLines, options).Trim();
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

        headingText = JoinParagraphLines(contentLines, options).Trim();
        return headingText.Length > 0;
    }

    private static string JoinParagraphLines(List<string> lines, MarkdownReaderOptions options) {
        var preservedInlineLineBreaks = FindParagraphLineBreaksInsideMatchedInlinePreserveSpans(lines, options);
        var sb = new StringBuilder();
        bool prevHard = false;
        for (int i = 0; i < lines.Count; i++) {
            var raw = lines[i] ?? string.Empty;
            var joinInfo = GetParagraphLineJoinInfo(
                raw,
                absoluteLine: 1,
                startColumn: 1,
                options,
                sourceTextMap: null,
                hasFollowingLine: i + 1 < lines.Count,
                preserveLineEndingInsideInlineSpan: i < preservedInlineLineBreaks.Length && preservedInlineLineBreaks[i]);

            if (i > 0) sb.Append(prevHard ? "\n" : " ");
            sb.Append(joinInfo.Text);
            prevHard = joinInfo.UsesLineBreakSeparator;
        }
        return sb.ToString();
    }

    private static (string Text, MarkdownInlineSourceMap? SourceMap) JoinParagraphLinesWithSourceMap(
        List<string> lines,
        int absoluteLineOffset,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        IReadOnlyList<int>? lineStartColumns = null) {
        if (state?.SourceTextMap == null) {
            var sourceText = JoinParagraphLines(lines, options);
            return (sourceText, null);
        }

        var textBuilder = new StringBuilder();
        var pointList = new List<MarkdownSourcePoint?>();
        var tokenSpanList = new List<MarkdownSourceSpan?>();
        var tokenLiteralList = new List<string?>();
        var preservedInlineLineBreaks = FindParagraphLineBreaksInsideMatchedInlinePreserveSpans(lines, options);
        ParagraphLineJoinInfo? previousJoinInfo = null;
        var previousAbsoluteLine = absoluteLineOffset + 1;
        var previousJoinColumn = 1;

        for (var i = 0; i < lines.Count; i++) {
            var absoluteLine = absoluteLineOffset + i + 1;
            var lineStartColumn = GetParagraphLineStartColumn(lineStartColumns, i);
            if (previousJoinInfo.HasValue) {
                var softLazyQuoteBreak = IsLazyQuoteContinuationLine(state, absoluteLineOffset + i) &&
                    !previousJoinInfo.Value.HardBreak;
                textBuilder.Append(previousJoinInfo.Value.UsesLineBreakSeparator || softLazyQuoteBreak ? '\n' : ' ');
                pointList.Add(state.SourceTextMap.CreatePoint(previousAbsoluteLine, previousJoinColumn));
                tokenSpanList.Add(previousJoinInfo.Value.HardBreak
                    ? previousJoinInfo.Value.HardBreakMarkerSpan
                    : softLazyQuoteBreak
                        ? CreateSpan(
                            state,
                            previousAbsoluteLine,
                            previousJoinColumn,
                            previousAbsoluteLine,
                            previousJoinColumn)
                        : null);
                tokenLiteralList.Add(previousJoinInfo.Value.HardBreak
                    ? previousJoinInfo.Value.HardBreakMarker
                    : softLazyQuoteBreak
                        ? "\n"
                        : null);
            }

            var raw = lines[i] ?? string.Empty;
            var joinInfo = GetParagraphLineJoinInfo(
                raw,
                absoluteLine,
                lineStartColumn,
                options,
                state.SourceTextMap,
                hasFollowingLine: i + 1 < lines.Count,
                preserveLineEndingInsideInlineSpan: i < preservedInlineLineBreaks.Length && preservedInlineLineBreaks[i]);
            textBuilder.Append(joinInfo.Text);
            var sourceColumn = lineStartColumn;
            for (var charIndex = 0; charIndex < joinInfo.Text.Length; charIndex++) {
                pointList.Add(state.SourceTextMap.CreatePoint(absoluteLine, sourceColumn));
                tokenSpanList.Add(null);
                tokenLiteralList.Add(null);
                sourceColumn = MarkdownSourceColumns.AdvanceColumn(sourceColumn, joinInfo.Text[charIndex]);
            }

            previousAbsoluteLine = absoluteLine;
            previousJoinColumn = Math.Max(lineStartColumn, sourceColumn - 1);
            previousJoinInfo = joinInfo;
        }

        var text = textBuilder.ToString();
        if (string.IsNullOrEmpty(text)) {
            return (text, null);
        }

        return (text, new MarkdownInlineSourceMap(pointList.ToArray(), tokenSpanList.ToArray(), tokenLiteralList.ToArray()));
    }

    private static int GetParagraphLineStartColumn(IReadOnlyList<int>? lineStartColumns, int lineIndex) {
        if (lineStartColumns == null
            || lineIndex < 0
            || lineIndex >= lineStartColumns.Count
            || lineStartColumns[lineIndex] <= 0) {
            return 1;
        }

        return lineStartColumns[lineIndex];
    }

    private static (string Text, MarkdownInlineSourceMap? SourceMap) JoinParagraphSourceLinesWithSourceMap(
        List<MarkdownSourceLineSlice> lines,
        MarkdownReaderOptions options,
        MarkdownReaderState? state) {
        if (lines == null || lines.Count == 0) {
            return (string.Empty, null);
        }

        var plainLines = new List<string>(lines.Count);
        for (int i = 0; i < lines.Count; i++) {
            plainLines.Add(lines[i].Text);
        }

        if (state?.SourceTextMap == null) {
            var sourceText = JoinParagraphSourceLines(lines, options);
            return (sourceText, null);
        }

        var textBuilder = new StringBuilder();
        var points = new List<MarkdownSourcePoint?>();
        var tokenSpans = new List<MarkdownSourceSpan?>();
        var tokenLiterals = new List<string?>();
        var preservedInlineLineBreaks = FindParagraphLineBreaksInsideMatchedInlinePreserveSpans(plainLines, options);
        MarkdownSourceLineSlice? previousLine = null;
        ParagraphLineJoinInfo? previousJoinInfo = null;

        for (var i = 0; i < lines.Count; i++) {
            var slice = lines[i];
            if (previousLine.HasValue && previousJoinInfo.HasValue) {
                var softLazyQuoteBreak = IsLazyQuoteSoftBreak(previousLine.Value, slice) &&
                    !previousJoinInfo.Value.HardBreak;
                textBuilder.Append(previousJoinInfo.Value.UsesLineBreakSeparator || softLazyQuoteBreak ? '\n' : ' ');
                var previousJoinColumn = GetEndColumn(previousLine.Value.StartColumn, previousJoinInfo.Value.Text);
                points.Add(state.SourceTextMap.CreatePoint(previousLine.Value.AbsoluteLine, previousJoinColumn));
                tokenSpans.Add(previousJoinInfo.Value.HardBreak
                    ? previousJoinInfo.Value.HardBreakMarkerSpan
                    : softLazyQuoteBreak
                        ? CreateSpan(
                            state,
                            previousLine.Value.AbsoluteLine,
                            previousJoinColumn,
                            previousLine.Value.AbsoluteLine,
                            previousJoinColumn)
                        : null);
                tokenLiterals.Add(previousJoinInfo.Value.HardBreak
                    ? previousJoinInfo.Value.HardBreakMarker
                    : softLazyQuoteBreak
                        ? "\n"
                        : null);
            }

            var joinInfo = GetParagraphLineJoinInfo(
                slice.Text,
                slice.AbsoluteLine,
                slice.StartColumn,
                options,
                state.SourceTextMap,
                hasFollowingLine: i + 1 < lines.Count,
                preserveLineEndingInsideInlineSpan: i < preservedInlineLineBreaks.Length && preservedInlineLineBreaks[i]);
            textBuilder.Append(joinInfo.Text);
            var sourceColumn = slice.StartColumn;
            for (var charIndex = 0; charIndex < joinInfo.Text.Length; charIndex++) {
                points.Add(state.SourceTextMap.CreatePoint(slice.AbsoluteLine, sourceColumn));
                tokenSpans.Add(null);
                tokenLiterals.Add(null);
                sourceColumn = MarkdownSourceColumns.AdvanceColumn(sourceColumn, joinInfo.Text[charIndex]);
            }

            previousLine = slice;
            previousJoinInfo = joinInfo;
        }

        var text = textBuilder.ToString();
        if (string.IsNullOrEmpty(text)) {
            return (text, null);
        }

        return (text, new MarkdownInlineSourceMap(points.ToArray(), tokenSpans.ToArray(), tokenLiterals.ToArray()));
    }

    private static string JoinParagraphSourceLines(List<MarkdownSourceLineSlice> lines, MarkdownReaderOptions options) {
        if (lines == null || lines.Count == 0) {
            return string.Empty;
        }

        var plainLines = new List<string>(lines.Count);
        for (int i = 0; i < lines.Count; i++) {
            plainLines.Add(lines[i].Text);
        }

        var preservedInlineLineBreaks = FindParagraphLineBreaksInsideMatchedInlinePreserveSpans(plainLines, options);
        var sb = new StringBuilder();
        ParagraphLineJoinInfo? previousJoinInfo = null;
        MarkdownSourceLineSlice? previousLine = null;
        for (int i = 0; i < lines.Count; i++) {
            var slice = lines[i];
            if (previousLine.HasValue && previousJoinInfo.HasValue) {
                sb.Append(previousJoinInfo.Value.UsesLineBreakSeparator ||
                          (IsLazyQuoteSoftBreak(previousLine.Value, slice) && !previousJoinInfo.Value.HardBreak)
                    ? '\n'
                    : ' ');
            }

            var joinInfo = GetParagraphLineJoinInfo(
                slice.Text,
                slice.AbsoluteLine,
                slice.StartColumn,
                options,
                sourceTextMap: null,
                hasFollowingLine: i + 1 < lines.Count,
                preserveLineEndingInsideInlineSpan: i < preservedInlineLineBreaks.Length && preservedInlineLineBreaks[i]);
            sb.Append(joinInfo.Text);
            previousLine = slice;
            previousJoinInfo = joinInfo;
        }

        return sb.ToString();
    }

    private static bool IsLazyQuoteSoftBreak(MarkdownSourceLineSlice previousLine, MarkdownSourceLineSlice currentLine) =>
        previousLine.IsLazyQuoteContinuation || currentLine.IsLazyQuoteContinuation;

    private static bool IsLazyQuoteContinuationLine(MarkdownReaderState? state, int zeroBasedLineIndex) =>
        state?.LazyQuoteContinuationLines.Contains(zeroBasedLineIndex) == true;

    private static MarkdownInlineSourceMap? BuildInlineSourceMapForSingleLine(
        string text,
        int absoluteLine,
        int startColumn,
        MarkdownReaderState? state) {
        if (state?.SourceTextMap == null || string.IsNullOrEmpty(text)) {
            return null;
        }

        var points = new MarkdownSourcePoint?[text.Length];
        var sourceColumn = startColumn;
        for (var i = 0; i < text.Length; i++) {
            points[i] = state.SourceTextMap.CreatePoint(absoluteLine, sourceColumn);
            sourceColumn = MarkdownSourceColumns.AdvanceColumn(sourceColumn, text[i]);
        }

        return new MarkdownInlineSourceMap(points);
    }

    private static int GetEndColumn(int startColumn, string? text) {
        if (string.IsNullOrEmpty(text)) {
            return startColumn;
        }

        var sourceColumn = startColumn;
        for (var i = 0; i < text!.Length; i++) {
            sourceColumn = MarkdownSourceColumns.AdvanceColumn(sourceColumn, text[i]);
        }

        return Math.Max(startColumn, sourceColumn - 1);
    }

    private static string ConsumeTrailingBackslashHardBreak(string trimmed, MarkdownReaderOptions options, out bool hardBreak) {
        hardBreak = false;
        if (options == null || !options.BackslashHardBreaks) return trimmed ?? string.Empty;
        if (string.IsNullOrEmpty(trimmed)) return string.Empty;
        if (trimmed[trimmed.Length - 1] != '\\') return trimmed;
        hardBreak = true;
        return trimmed.Substring(0, trimmed.Length - 1);
    }

    private static ParagraphLineJoinInfo GetParagraphLineJoinInfo(
        string raw,
        int absoluteLine,
        int startColumn,
        MarkdownReaderOptions options,
        MarkdownSourceTextMap? sourceTextMap,
        bool hasFollowingLine,
        bool preserveLineEndingInsideInlineSpan = false) {
        raw ??= string.Empty;
        if (preserveLineEndingInsideInlineSpan) {
            return new ParagraphLineJoinInfo(raw, hardBreak: false, hardBreakMarker: null, hardBreakMarkerSpan: null, preserveLineBreak: true);
        }

        bool spaceHardBreak = EndsWithTwoSpacesLine(raw);
        var trimmed = raw.TrimEnd();
        bool backslashHardBreak = false;
        var text = hasFollowingLine
            ? ConsumeTrailingBackslashHardBreak(trimmed, options, out backslashHardBreak)
            : trimmed;

        if (backslashHardBreak) {
            var markerColumn = AdvanceParagraphSourceColumn(startColumn, trimmed, Math.Max(0, trimmed.Length - 1));
            return new ParagraphLineJoinInfo(
                text,
                hardBreak: true,
                hardBreakMarker: "\\",
                hardBreakMarkerSpan: sourceTextMap?.CreateSpan(absoluteLine, markerColumn, absoluteLine, markerColumn));
        }

        if (spaceHardBreak) {
            var markerStartIndex = raw.TrimEnd(' ').Length;
            var markerLength = raw.Length - markerStartIndex;
            var markerStartColumn = AdvanceParagraphSourceColumn(startColumn, raw, markerStartIndex);
            var markerEndColumn = AdvanceParagraphSourceColumn(startColumn, raw, raw.Length) - 1;
            return new ParagraphLineJoinInfo(
                text,
                hardBreak: true,
                hardBreakMarker: markerLength > 0 ? raw.Substring(markerStartIndex, markerLength) : string.Empty,
                hardBreakMarkerSpan: sourceTextMap?.CreateSpan(absoluteLine, markerStartColumn, absoluteLine, markerEndColumn));
        }

        if (hasFollowingLine && options?.SoftLineBreaksAsHardLineBreaks == true) {
            return new ParagraphLineJoinInfo(text, hardBreak: true, hardBreakMarker: null, hardBreakMarkerSpan: null);
        }

        return new ParagraphLineJoinInfo(text, hardBreak: false, hardBreakMarker: null, hardBreakMarkerSpan: null);
    }

    private static int AdvanceParagraphSourceColumn(int startColumn, string? text, int endExclusive) {
        var column = Math.Max(1, startColumn);
        var value = text ?? string.Empty;
        if (value.Length == 0 || endExclusive <= 0) {
            return column;
        }

        int limit = Math.Min(value.Length, endExclusive);
        for (int i = 0; i < limit; i++) {
            column = MarkdownSourceColumns.AdvanceColumn(column, value[i]);
        }

        return column;
    }

}

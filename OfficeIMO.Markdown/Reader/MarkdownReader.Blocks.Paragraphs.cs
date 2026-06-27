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
        public MarkdownSourceLineSlice(string text, int absoluteLine, int startColumn, bool isLazyQuoteContinuation = false) {
            Text = text ?? string.Empty;
            AbsoluteLine = absoluteLine;
            StartColumn = startColumn < 1 ? 1 : startColumn;
            IsLazyQuoteContinuation = isLazyQuoteContinuation;
        }

        public string Text { get; }
        public int AbsoluteLine { get; }
        public int StartColumn { get; }
        public bool IsLazyQuoteContinuation { get; }
    }

    private readonly struct ParagraphLineJoinInfo {
        public ParagraphLineJoinInfo(
            string text,
            bool hardBreak,
            string? hardBreakMarker,
            MarkdownSourceSpan? hardBreakMarkerSpan) {
            Text = text ?? string.Empty;
            HardBreak = hardBreak;
            HardBreakMarker = hardBreakMarker;
            HardBreakMarkerSpan = hardBreakMarkerSpan;
        }

        public string Text { get; }
        public bool HardBreak { get; }
        public string? HardBreakMarker { get; }
        public MarkdownSourceSpan? HardBreakMarkerSpan { get; }
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
        if (item.SyntaxChildren.Count > 0) return;
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

        var lastLine = slices[slices.Count - 1].Text;
        var localSpan = new MarkdownSourceSpan(1, 1, slices.Count, Math.Max(1, lastLine.Length));
        var localNode = BuildSyntaxNode(block, localSpan);
        var remappedNode = RemapNestedSyntaxNode(slices, localNode);
        SynchronizeOwnedSyntaxCaches(remappedNode);
        MarkdownObjectTreeBinder.BindSourceSpans(remappedNode);
        item.SyntaxChildren.Add(remappedNode);
    }

    private static List<MarkdownSourceLineSlice> BuildListItemNestedSourceLines(
        string[] sourceLines,
        int continuationIndent,
        int startLineIndex,
        int endExclusiveLineIndex,
        MarkdownReaderState? state) {

        var count = Math.Max(0, Math.Min(endExclusiveLineIndex, sourceLines.Length) - Math.Max(0, startLineIndex));
        var slices = new List<MarkdownSourceLineSlice>(count);
        var lineOffset = state?.SourceLineOffset ?? 0;
        var start = Math.Max(0, startLineIndex);
        var end = Math.Min(endExclusiveLineIndex, sourceLines.Length);

        for (var i = start; i < end; i++) {
            var line = sourceLines[i] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(line)) {
                slices.Add(new MarkdownSourceLineSlice(string.Empty, lineOffset + i + 1, 1));
                continue;
            }

            if (CountLeadingIndentColumns(line) >= continuationIndent) {
                slices.Add(new MarkdownSourceLineSlice(
                    StripLeadingIndentColumns(line, continuationIndent),
                    lineOffset + i + 1,
                    continuationIndent + 1));
                continue;
            }

            var leadingColumns = CountLeadingIndentColumns(line);
            slices.Add(new MarkdownSourceLineSlice(
                line.TrimStart(),
                lineOffset + i + 1,
                leadingColumns + 1));
        }

        return slices;
    }

    private static ListItem CreateListItemFromLeadLines(List<string> lines, bool isTask, bool done, MarkdownReaderOptions options, MarkdownReaderState? state, List<MarkdownSourceLineSlice>? sourceLines = null) {
        if (TryCreateListItemFromLeadBlocks(lines, isTask, done, options, state, sourceLines, out var blockLeadItem)) {
            return blockLeadItem;
        }

        if (TryParseListItemLeadSetextBlocks(lines, options, state, out var leadBlocks)) {
            var headingItem = isTask ? ListItem.TaskInlines(new InlineSequence(), done) : new ListItem(new InlineSequence());
            for (int i = 0; i < leadBlocks.Count; i++) {
                headingItem.Children.Add(leadBlocks[i]);
            }
            return headingItem;
        }

        int firstBlank = lines.FindIndex(string.IsNullOrWhiteSpace);
        if (firstBlank <= 0) {
            var paragraphs = sourceLines != null && sourceLines.Count == lines.Count
                ? ParseParagraphsFromSourceLines(sourceLines, options, state)
                : ParseParagraphsFromLines(lines, options, state);
            var item = isTask ? ListItem.TaskInlines(paragraphs[0], done) : new ListItem(paragraphs[0]);
            for (int i = 1; i < paragraphs.Count; i++) {
                item.AdditionalParagraphs.Add(paragraphs[i]);
            }
            return item;
        }

        var firstParagraph = sourceLines != null && sourceLines.Count >= firstBlank
            ? ParseParagraphsFromSourceLines(sourceLines.GetRange(0, firstBlank), options, state)[0]
            : ParseParagraphsFromLines(lines.GetRange(0, firstBlank), options, state)[0];
        var mixedItem = isTask ? ListItem.TaskInlines(firstParagraph, done) : new ListItem(firstParagraph);

        if (firstBlank + 1 >= lines.Count) return mixedItem;

        var trailingLines = lines.GetRange(firstBlank + 1, lines.Count - firstBlank - 1);
        if (trailingLines.TrueForAll(string.IsNullOrWhiteSpace)) return mixedItem;

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
                    return mixedItem;
                }

                for (int i = 0; i < trailingBlocksFromSource.Count; i++) {
                    mixedItem.Children.Add(trailingBlocksFromSource[i]);
                }
                mixedItem.ForceLoose = true;
                return mixedItem;
            }
        }

        var trailingBlocks = ParseBlocksFromLines(trailingLines.ToArray(), options, state ?? new MarkdownReaderState());
        if (mixedItem.TryAbsorbTrailingParagraphBlocks(trailingBlocks)) return mixedItem;

        for (int i = 0; i < trailingBlocks.Count; i++) {
            mixedItem.Children.Add(trailingBlocks[i]);
        }
        mixedItem.ForceLoose = true;
        return mixedItem;
    }

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
            blockLeadItem.Children.Add(leadBlocks[i]);
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

        if (options.OrderedLists && IsOrderedListLine(trimmed, out _, out _)) {
            return true;
        }

        return false;
    }

    private static bool TryParseListItemLeadSetextBlocks(List<string> lines, MarkdownReaderOptions options, MarkdownReaderState? state, out List<IMarkdownBlock> blocks) {
        blocks = new List<IMarkdownBlock>();
        if (lines == null || lines.Count == 0 || options == null || !options.Headings) return false;

        if (!TryGetLeadingSetextHeadingPrefix(lines, options, out int headingLineCount, out int level, out string headingText)) return false;

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

        for (int prefixLength = 2; prefixLength <= maxPrefixLength; prefixLength++) {
            var candidate = lines.GetRange(0, prefixLength);
            if (!TryParseSetextHeadingParagraphLines(candidate, options, out level, out headingText)) continue;

            headingLineCount = prefixLength;
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
        var sb = new StringBuilder();
        bool prevHard = false;
        for (int i = 0; i < lines.Count; i++) {
            var raw = lines[i] ?? string.Empty;
            bool hard = EndsWithTwoSpacesLine(raw);
            var trimmed = raw.TrimEnd();
            trimmed = ConsumeTrailingBackslashHardBreak(trimmed, options, out bool slashHard);
            hard = hard || slashHard;

            if (i > 0) sb.Append(prevHard ? "\n" : " ");
            sb.Append(trimmed);
            prevHard = hard;
        }
        return sb.ToString();
    }

    private static (string Text, MarkdownInlineSourceMap? SourceMap) JoinParagraphLinesWithSourceMap(
        List<string> lines,
        int absoluteLineOffset,
        MarkdownReaderOptions options,
        MarkdownReaderState? state) {
        var text = JoinParagraphLines(lines, options);
        if (state?.SourceTextMap == null || string.IsNullOrEmpty(text)) {
            return (text, null);
        }

        var points = new MarkdownSourcePoint?[text.Length];
        var tokenSpans = new MarkdownSourceSpan?[text.Length];
        var tokenLiterals = new string?[text.Length];
        var cursor = 0;
        var previousLineForJoin = absoluteLineOffset + 1;
        var previousJoinColumn = 1;
        MarkdownSourceSpan? previousHardBreakMarkerSpan = null;
        string? previousHardBreakMarker = null;

        for (var i = 0; i < lines.Count; i++) {
            var raw = lines[i] ?? string.Empty;
            var absoluteLine = absoluteLineOffset + i + 1;
            var joinInfo = GetParagraphLineJoinInfo(raw, absoluteLine, 1, options, state.SourceTextMap);

            if (i > 0 && cursor < points.Length) {
                points[cursor] = state.SourceTextMap.CreatePoint(previousLineForJoin, previousJoinColumn);
                tokenSpans[cursor] = previousHardBreakMarkerSpan;
                tokenLiterals[cursor] = previousHardBreakMarker;
                cursor++;
            }

            for (var charIndex = 0; charIndex < joinInfo.Text.Length && cursor < points.Length; charIndex++) {
                points[cursor++] = state.SourceTextMap.CreatePoint(absoluteLine, charIndex + 1);
            }

            previousLineForJoin = absoluteLine;
            previousJoinColumn = Math.Max(1, joinInfo.Text.Length);
            previousHardBreakMarkerSpan = joinInfo.HardBreak ? joinInfo.HardBreakMarkerSpan : null;
            previousHardBreakMarker = joinInfo.HardBreak ? joinInfo.HardBreakMarker : null;
        }

        if (cursor < points.Length) {
            Array.Resize(ref points, cursor);
            Array.Resize(ref tokenSpans, cursor);
            Array.Resize(ref tokenLiterals, cursor);
        }

        return (text, new MarkdownInlineSourceMap(points, tokenSpans, tokenLiterals));
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

        var text = JoinParagraphLines(plainLines, options);
        if (state?.SourceTextMap == null || string.IsNullOrEmpty(text)) {
            return (text, null);
        }

        var points = new MarkdownSourcePoint?[text.Length];
        var tokenSpans = new MarkdownSourceSpan?[text.Length];
        var tokenLiterals = new string?[text.Length];
        var cursor = 0;
        var previousLine = lines[0].AbsoluteLine;
        var previousJoinColumn = lines[0].StartColumn;
        MarkdownSourceSpan? previousHardBreakMarkerSpan = null;
        string? previousHardBreakMarker = null;

        for (var i = 0; i < lines.Count; i++) {
            if (i > 0 && cursor < points.Length) {
                points[cursor] = state.SourceTextMap.CreatePoint(previousLine, previousJoinColumn);
                tokenSpans[cursor] = previousHardBreakMarkerSpan;
                tokenLiterals[cursor] = previousHardBreakMarker;
                cursor++;
            }

            var slice = lines[i];
            var joinInfo = GetParagraphLineJoinInfo(slice.Text, slice.AbsoluteLine, slice.StartColumn, options, state.SourceTextMap);
            for (var charIndex = 0; charIndex < joinInfo.Text.Length && cursor < points.Length; charIndex++) {
                points[cursor++] = state.SourceTextMap.CreatePoint(slice.AbsoluteLine, slice.StartColumn + charIndex);
            }

            previousLine = slice.AbsoluteLine;
            previousJoinColumn = slice.StartColumn + Math.Max(0, joinInfo.Text.Length - 1);
            previousHardBreakMarkerSpan = joinInfo.HardBreak ? joinInfo.HardBreakMarkerSpan : null;
            previousHardBreakMarker = joinInfo.HardBreak ? joinInfo.HardBreakMarker : null;
        }

        if (cursor < points.Length) {
            Array.Resize(ref points, cursor);
            Array.Resize(ref tokenSpans, cursor);
            Array.Resize(ref tokenLiterals, cursor);
        }

        return (text, new MarkdownInlineSourceMap(points, tokenSpans, tokenLiterals));
    }

    private static MarkdownInlineSourceMap? BuildInlineSourceMapForSingleLine(
        string text,
        int absoluteLine,
        int startColumn,
        MarkdownReaderState? state) {
        if (state?.SourceTextMap == null || string.IsNullOrEmpty(text)) {
            return null;
        }

        var points = new MarkdownSourcePoint?[text.Length];
        for (var i = 0; i < text.Length; i++) {
            points[i] = state.SourceTextMap.CreatePoint(absoluteLine, startColumn + i);
        }

        return new MarkdownInlineSourceMap(points);
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
        MarkdownSourceTextMap? sourceTextMap) {
        raw ??= string.Empty;

        bool spaceHardBreak = EndsWithTwoSpacesLine(raw);
        var trimmed = raw.TrimEnd();
        var text = ConsumeTrailingBackslashHardBreak(trimmed, options, out bool backslashHardBreak);

        if (backslashHardBreak) {
            var markerColumn = startColumn + Math.Max(0, trimmed.Length - 1);
            return new ParagraphLineJoinInfo(
                text,
                hardBreak: true,
                hardBreakMarker: "\\",
                hardBreakMarkerSpan: sourceTextMap?.CreateSpan(absoluteLine, markerColumn, absoluteLine, markerColumn));
        }

        if (spaceHardBreak) {
            var markerStartIndex = raw.TrimEnd(' ').Length;
            var markerLength = raw.Length - markerStartIndex;
            var markerStartColumn = startColumn + markerStartIndex;
            var markerEndColumn = startColumn + raw.Length - 1;
            return new ParagraphLineJoinInfo(
                text,
                hardBreak: true,
                hardBreakMarker: markerLength > 0 ? raw.Substring(markerStartIndex, markerLength) : string.Empty,
                hardBreakMarkerSpan: sourceTextMap?.CreateSpan(absoluteLine, markerStartColumn, absoluteLine, markerEndColumn));
        }

        return new ParagraphLineJoinInfo(text, hardBreak: false, hardBreakMarker: null, hardBreakMarkerSpan: null);
    }

}

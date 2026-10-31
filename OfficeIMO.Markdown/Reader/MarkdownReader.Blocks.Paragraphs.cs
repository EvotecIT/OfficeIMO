using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

internal static partial class MarkdownReader {
    private static List<IMarkdownBlock> ParseBlocksFromLines(string[] lines, MarkdownLoadOptions options, MarkdownReaderState state, List<MarkdownSyntaxNode>? syntaxNodes = null, int lineOffset = 0) {
        var doc = MarkdownDocument.Create();
        var opt = CloneOptionsWithoutFrontMatter(options);
        var pipeline = MarkdownReaderPipeline.Default(opt);
        int previousLineOffset = state.SourceLineOffset;
        bool previousSyntheticTerminalLine = state.HasSyntheticTerminalLine;
        state.SourceLineOffset = lineOffset;
        state.HasSyntheticTerminalLine = false;

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
            state.HasSyntheticTerminalLine = previousSyntheticTerminalLine;
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

    private static List<InlineSequence> ParseParagraphsFromLines(List<string> lines, MarkdownLoadOptions options, MarkdownReaderState? state) {
        var paragraphs = new List<InlineSequence>();
        if (lines == null || lines.Count == 0) {
            paragraphs.Add(ParseInlines(string.Empty, options, state));
            return paragraphs;
        }

        if (lines.Count == 1 && !string.IsNullOrEmpty(lines[0])) {
            paragraphs.Add(ParseJoinedParagraphLines(lines, options, state));
            return paragraphs;
        }

        var cur = new List<string>();
        for (int i = 0; i < lines.Count; i++) {
            var ln = lines[i] ?? string.Empty;
            if (ln.Length == 0) {
                if (cur.Count > 0) {
                    paragraphs.Add(ParseJoinedParagraphLines(cur, options, state));
                    cur.Clear();
                }
                continue;
            }
            cur.Add(ln);
        }
        if (cur.Count > 0) paragraphs.Add(ParseJoinedParagraphLines(cur, options, state));

        if (paragraphs.Count == 0) paragraphs.Add(ParseInlines(string.Empty, options, state));
        return paragraphs;
    }

    private static List<InlineSequence> ParseParagraphsFromSourceLines(List<MarkdownSourceLineSlice> lines, MarkdownLoadOptions options, MarkdownReaderState? state) {
        var paragraphs = new List<InlineSequence>();
        if (lines == null || lines.Count == 0) {
            paragraphs.Add(ParseInlines(string.Empty, options, state));
            return paragraphs;
        }

        if (lines.Count == 1 && !string.IsNullOrEmpty(lines[0].Text)) {
            paragraphs.Add(ParseParagraphFromSourceLines(lines, options, state));
            return paragraphs;
        }

        var current = new List<MarkdownSourceLineSlice>();
        for (int i = 0; i < lines.Count; i++) {
            if (string.IsNullOrEmpty(lines[i].Text)) {
                if (current.Count > 0) {
                    paragraphs.Add(ParseParagraphFromSourceLines(current, options, state));
                    current.Clear();
                }
                continue;
            }

            current.Add(lines[i]);
        }

        if (current.Count > 0) {
            paragraphs.Add(ParseParagraphFromSourceLines(current, options, state));
        }

        if (paragraphs.Count == 0) {
            paragraphs.Add(ParseInlines(string.Empty, options, state));
        }

        return paragraphs;
    }

    private static InlineSequence ParseParagraphFromSourceLines(
        List<MarkdownSourceLineSlice> lines,
        MarkdownLoadOptions options,
        MarkdownReaderState? state) {
        var (text, sourceMap) = JoinParagraphSourceLinesWithSourceMap(lines, options, state);
        return ParseInlines(text, options, state, sourceMap);
    }

    private static List<ParagraphBlock> ParseParagraphBlocksFromLines(List<string> lines, MarkdownLoadOptions options, MarkdownReaderState? state) {
        var paragraphInlines = ParseParagraphsFromLines(lines, options, state);
        var blocks = new List<ParagraphBlock>(paragraphInlines.Count);
        for (int i = 0; i < paragraphInlines.Count; i++) {
            blocks.Add(new ParagraphBlock(paragraphInlines[i]));
        }
        return blocks;
    }

    private static List<ParagraphBlock> ParseParagraphBlocksFromSourceLines(List<MarkdownSourceLineSlice> lines, MarkdownLoadOptions options, MarkdownReaderState? state) {
        var paragraphInlines = ParseParagraphsFromSourceLines(lines, options, state);
        var blocks = new List<ParagraphBlock>(paragraphInlines.Count);
        for (int i = 0; i < paragraphInlines.Count; i++) {
            blocks.Add(new ParagraphBlock(paragraphInlines[i]));
        }
        return blocks;
    }

    private static void AddListItemLeadSyntaxNodes(ListItem item, List<string> lines, int lineOffset, MarkdownLoadOptions options, MarkdownReaderState? state, List<MarkdownSourceLineSlice>? sourceLines = null) {
        if (item == null || lines == null || lines.Count == 0) return;
        if (HasNonSyntaxOnlyListItemLeadSyntaxChildren(item.SyntaxChildren)) return;
        int absoluteLineOffset = (state?.SourceLineOffset ?? 0) + lineOffset;

        if (lines.Count == 1 &&
            sourceLines?.Count == 1 &&
            item.Content.Nodes.Count > 0 &&
            !StartsListItemLeadWithStandaloneBlock(lines, options)) {
            var sourceLine = sourceLines[0];
            item.SyntaxChildren.Add(BuildSyntaxNode(item.LeadParagraphBlock, CreateSpan(
                state,
                sourceLine.AbsoluteLine,
                sourceLine.StartColumn,
                sourceLine.AbsoluteLine,
                sourceLine.StartColumn + Math.Max(0, sourceLine.Text.Length - 1))));
            return;
        }

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

        if (TryGetLeadingSetextHeadingPrefix(lines, options, out int headingLineCount, out level, out headingText, out _)) {
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

    private static void AddParagraphSyntaxNodes(List<MarkdownSyntaxNode> nodes, List<string> lines, int lineOffset, MarkdownLoadOptions options, MarkdownReaderState? state) {
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

    private static void AddParagraphSyntaxNodes(List<MarkdownSyntaxNode> nodes, List<MarkdownSourceLineSlice> lines, MarkdownLoadOptions options, MarkdownReaderState? state) {
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
        MarkdownLoadOptions options,
        MarkdownReaderState? state) {
        if (lines == null || lines.Count == 0) {
            return Array.Empty<MarkdownSyntaxNode>();
        }

        var effectiveState = state ?? new MarkdownReaderState();
        var (_, syntaxChildren) = ParseNestedMarkdownBlocks(lines, options, effectiveState);
        return syntaxChildren;
    }

    private static void FlushParagraphSyntaxNode(List<MarkdownSyntaxNode> nodes, List<string> lines, int startIndex, int endIndex, int lineOffset, MarkdownLoadOptions options, MarkdownReaderState? state) {
        if (nodes == null || lines == null || lines.Count == 0 || startIndex < 0 || endIndex < startIndex) return;

        var inlines = ParseJoinedParagraphLines(lines, options, state);
        var paragraph = new ParagraphBlock(inlines);
        nodes.Add(BuildSyntaxNode(paragraph, CreateLineSpan(state, lineOffset + startIndex + 1, lineOffset + endIndex + 1)));
    }

    private static void FlushParagraphSyntaxNode(List<MarkdownSyntaxNode> nodes, List<MarkdownSourceLineSlice> lines, MarkdownLoadOptions options, MarkdownReaderState? state) {
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
        if (item == null || block == null || state?.CaptureSyntaxTree != true) return;
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

        if (item == null || block == null || sourceLines == null || state?.CaptureSyntaxTree != true) return;

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

}

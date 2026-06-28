namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private const int DefinitionContinuationIndentOffset = 2;

    internal sealed class DefinitionListParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.DefinitionLists) return false;
            if (TryParseMarkdigDefinitionList(lines, ref i, options, doc, state)) return true;
            if (!ShouldTreatAsDefinitionLine(lines, i, options)) return false;
            var dl = new DefinitionListBlock();
            dl.SetParsingContext(options, state);
            int j = i;
            while (j < lines.Length && ShouldTreatAsDefinitionLine(lines, j, options)) {
                if (!TryGetDefinitionSeparator(lines[j], out var idx)) break;
                var line = lines[j] ?? string.Empty;
                int lineNumber = state.SourceLineOffset + j + 1;

                int termStartIndex = 0;
                while (termStartIndex < idx && char.IsWhiteSpace(line[termStartIndex])) {
                    termStartIndex++;
                }

                int termEndExclusive = idx;
                while (termEndExclusive > termStartIndex && char.IsWhiteSpace(line[termEndExclusive - 1])) {
                    termEndExclusive--;
                }

                int definitionStartIndex = idx + 1;
                while (definitionStartIndex < line.Length && char.IsWhiteSpace(line[definitionStartIndex])) {
                    definitionStartIndex++;
                }

                int definitionEndExclusive = line.Length;
                while (definitionEndExclusive > definitionStartIndex && char.IsWhiteSpace(line[definitionEndExclusive - 1])) {
                    definitionEndExclusive--;
                }

                var term = termStartIndex < termEndExclusive
                    ? line.Substring(termStartIndex, termEndExclusive - termStartIndex)
                    : string.Empty;
                var def = definitionStartIndex < definitionEndExclusive
                    ? line.Substring(definitionStartIndex, definitionEndExclusive - definitionStartIndex)
                    : string.Empty;

                var termSourceMap = BuildInlineSourceMapForSingleLine(term, lineNumber, termStartIndex + 1, state);
                var definitionSourceMap = BuildInlineSourceMapForSingleLine(def, lineNumber, definitionStartIndex + 1, state);
                var termInlines = ParseInlines(term, options, state, termSourceMap);
                var definitionInlines = ParseInlines(def, options, state, definitionSourceMap);
                var termSpan = CreateSpan(
                    state,
                    lineNumber,
                    termStartIndex + 1,
                    lineNumber,
                    Math.Max(termStartIndex + 1, termEndExclusive));
                var definitionSpan = CreateSpan(
                    state,
                    lineNumber,
                    definitionStartIndex + 1,
                    lineNumber,
                    Math.Max(definitionStartIndex + 1, definitionEndExclusive));
                var definitionSourceLines = new List<MarkdownSourceLineSlice> {
                    new MarkdownSourceLineSlice(def, lineNumber, definitionStartIndex + 1)
                };
                int next = j + 1;
                ConsumeDefinitionContinuationLines(
                    lines,
                    ref next,
                    CountLeadingIndentColumns(line) + DefinitionContinuationIndentOffset,
                    state.SourceLineOffset,
                    definitionSourceLines,
                    allowLazyContinuation: false);
                var (definitionBlocks, definitionSyntaxChildren) = ParseDefinitionBody(definitionSourceLines, options, state);
                if (definitionSyntaxChildren.Count > 0) {
                    definitionSpan = MarkdownBlockSyntaxBuilder.GetAggregateSpan(definitionSyntaxChildren) ?? definitionSpan;
                }

                var definitionObject = new DefinitionListDefinition(definitionBlocks);
                var group = new DefinitionListGroup(
                    new[] { termInlines },
                    new[] { definitionObject });
                var termObject = group.TermItems[0];
                var termNode = MarkdownBlockSyntaxBuilder.BuildInlineContainerNode(
                    MarkdownSyntaxKind.DefinitionTerm,
                    termInlines,
                    termSpan,
                    term,
                    associatedObject: termObject);
                var valueNode = new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.DefinitionValue,
                    definitionSpan,
                    RenderDefinitionLiteral(definitionBlocks, def),
                    definitionSyntaxChildren,
                    associatedObject: definitionObject);
                dl.AddParsedGroup(
                    group,
                    new MarkdownSyntaxNode(
                        MarkdownSyntaxKind.DefinitionGroup,
                        MarkdownBlockSyntaxBuilder.GetAggregateSpan(new[] { termNode, valueNode }),
                        children: new[] {
                            termNode,
                            valueNode
                        },
                        associatedObject: group));
                j = next;
                int separator = j;
                while (separator < lines.Length && string.IsNullOrWhiteSpace(lines[separator])) {
                    separator++;
                }

                if (separator < lines.Length && ShouldTreatAsDefinitionLine(lines, separator, options)) {
                    j = separator;
                }
            }
            doc.Add(dl); i = j; return true;
        }
    }

    private static bool TryParseMarkdigDefinitionList(
        string[] lines,
        ref int i,
        MarkdownReaderOptions options,
        MarkdownDoc doc,
        MarkdownReaderState state) {
        if (lines == null || i < 0 || i >= lines.Length) {
            return false;
        }

        if (!TryCollectMarkdigDefinitionTerms(lines, i, out var markerIndex, out var terms)) {
            return false;
        }

        var dl = new DefinitionListBlock();
        dl.SetParsingContext(options, state);
        int cursor = i;
        while (cursor < lines.Length &&
               TryCollectMarkdigDefinitionTerms(lines, cursor, out markerIndex, out terms)) {
            var definitions = new List<DefinitionListDefinition>();
            var definitionNodes = new List<MarkdownSyntaxNode>();
            cursor = markerIndex;

            while (cursor < lines.Length &&
                   TryGetMarkdigDefinitionMarker(lines[cursor], out int markerIndent, out int definitionStartIndex)) {
                var definition = ParseMarkdigMarkerDefinition(
                    lines,
                    ref cursor,
                    markerIndent,
                    definitionStartIndex,
                    options,
                    state,
                    out var definitionNode);

                if (definition != null && definitionNode != null) {
                    definitions.Add(definition);
                    definitionNodes.Add(definitionNode);
                }

                bool skippedBlankBeforeNextDefinition = false;
                int separator = cursor;
                while (separator < lines.Length && string.IsNullOrWhiteSpace(lines[separator])) {
                    skippedBlankBeforeNextDefinition = true;
                    separator++;
                }

                if (separator < lines.Length &&
                    TryGetMarkdigDefinitionMarker(lines[separator], out _, out _)) {
                    if (skippedBlankBeforeNextDefinition && definitions.Count > 0) {
                        definitions[definitions.Count - 1].ForceParagraphHtml = true;
                    }

                    cursor = separator;
                    continue;
                }

                break;
            }

            for (int termIndex = 0; termIndex < terms.Count; termIndex++) {
                terms[termIndex].Bind(options, state);
            }

            var group = new DefinitionListGroup(
                terms.Select(static term => term.Inlines),
                definitions);
            var groupChildren = new List<MarkdownSyntaxNode>(terms.Count + definitionNodes.Count);
            for (int termIndex = 0; termIndex < terms.Count; termIndex++) {
                var termObject = group.TermItems[termIndex];
                groupChildren.Add(MarkdownBlockSyntaxBuilder.BuildInlineContainerNode(
                    MarkdownSyntaxKind.DefinitionTerm,
                    terms[termIndex].Inlines,
                    terms[termIndex].Span,
                    terms[termIndex].Literal,
                    associatedObject: termObject));
            }

            groupChildren.AddRange(definitionNodes);
            dl.AddParsedGroup(
                group,
                new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.DefinitionGroup,
                    MarkdownBlockSyntaxBuilder.GetAggregateSpan(groupChildren),
                    children: groupChildren,
                    associatedObject: group));

            int nextGroup = cursor;
            while (nextGroup < lines.Length && string.IsNullOrWhiteSpace(lines[nextGroup])) {
                nextGroup++;
            }

            if (nextGroup < lines.Length &&
                TryCollectMarkdigDefinitionTerms(lines, nextGroup, out _, out _)) {
                cursor = nextGroup;
                continue;
            }

            break;
        }

        doc.Add(dl);
        i = cursor;
        return true;
    }

    private static bool TryCollectMarkdigDefinitionTerms(
        string[] lines,
        int startIndex,
        out int markerIndex,
        out List<ParsedDefinitionTermLine> terms) {
        markerIndex = -1;
        terms = new List<ParsedDefinitionTermLine>();
        int cursor = startIndex;
        while (cursor < lines.Length) {
            var line = lines[cursor] ?? string.Empty;
            if (TryGetMarkdigDefinitionMarker(line, out _, out _)) {
                markerIndex = cursor;
                return terms.Count > 0;
            }

            if (!IsMarkdigDefinitionTermCandidate(line)) {
                return false;
            }

            terms.Add(CreateMarkdigDefinitionTerm(line, cursor));
            cursor++;
        }

        return false;
    }

    private static ParsedDefinitionTermLine CreateMarkdigDefinitionTerm(string line, int zeroBasedLineIndex) {
        int termStartIndex = 0;
        while (termStartIndex < line.Length && char.IsWhiteSpace(line[termStartIndex])) {
            termStartIndex++;
        }

        int termEndExclusive = line.Length;
        while (termEndExclusive > termStartIndex && char.IsWhiteSpace(line[termEndExclusive - 1])) {
            termEndExclusive--;
        }

        return new ParsedDefinitionTermLine(
            zeroBasedLineIndex,
            termStartIndex,
            termEndExclusive,
            termEndExclusive > termStartIndex
                ? line.Substring(termStartIndex, termEndExclusive - termStartIndex)
                : string.Empty);
    }

    private static bool IsMarkdigDefinitionTermCandidate(string line) {
        if (string.IsNullOrWhiteSpace(line)) return false;
        if (CountLeadingIndentColumns(line) >= 4) return false;
        var trimmed = line.TrimStart();
        if (IsAtxHeading(trimmed, out _, out _)) return false;
        if (IsUnorderedListLine(trimmed, out _, out _, out _)) return false;
        if (IsOrderedListLine(trimmed, out _, out _)) return false;
        if (StartsWithReferenceDefinitionLikeLabel(trimmed)) return false;
        return true;
    }

    private static bool TryGetMarkdigDefinitionMarker(
        string line,
        out int markerIndent,
        out int definitionStartIndex) {
        markerIndent = 0;
        definitionStartIndex = -1;
        if (string.IsNullOrEmpty(line)) return false;

        int rawIndex = 0;
        while (rawIndex < line.Length && line[rawIndex] == ' ') {
            markerIndent++;
            rawIndex++;
        }

        if (markerIndent >= 4 || rawIndex >= line.Length || line[rawIndex] != ':') {
            return false;
        }

        rawIndex++;
        int spacesAfterMarker = 0;
        while (rawIndex < line.Length && line[rawIndex] == ' ') {
            spacesAfterMarker++;
            rawIndex++;
        }

        if (rawIndex < line.Length && line[rawIndex] == '\t') {
            definitionStartIndex = rawIndex + 1;
            return true;
        }

        if (spacesAfterMarker >= 3) {
            definitionStartIndex = rawIndex;
            return true;
        }

        return false;
    }

    private static DefinitionListDefinition? ParseMarkdigMarkerDefinition(
        string[] lines,
        ref int index,
        int markerIndent,
        int definitionStartIndex,
        MarkdownReaderOptions options,
        MarkdownReaderState state,
        out MarkdownSyntaxNode? definitionNode) {
        var line = lines[index] ?? string.Empty;
        int lineNumber = state.SourceLineOffset + index + 1;
        int definitionEndExclusive = line.Length;
        while (definitionEndExclusive > definitionStartIndex && char.IsWhiteSpace(line[definitionEndExclusive - 1])) {
            definitionEndExclusive--;
        }

        var definitionLiteral = definitionStartIndex < definitionEndExclusive
            ? line.Substring(definitionStartIndex, definitionEndExclusive - definitionStartIndex)
            : string.Empty;
        var definitionSourceLines = new List<MarkdownSourceLineSlice>();
        if (!string.IsNullOrEmpty(definitionLiteral)) {
            definitionSourceLines.Add(new MarkdownSourceLineSlice(definitionLiteral, lineNumber, definitionStartIndex + 1));
        }

        int next = index + 1;
        ConsumeDefinitionContinuationLines(
            lines,
            ref next,
            markerIndent + DefinitionContinuationIndentOffset,
            state.SourceLineOffset,
            definitionSourceLines,
            allowLazyContinuation: true);
        index = next;

        if (definitionSourceLines.Count == 0) {
            definitionNode = null;
            return null;
        }

        var (definitionBlocks, definitionSyntaxChildren) = ParseDefinitionBody(definitionSourceLines, options, state);
        var definitionSpan = MarkdownBlockSyntaxBuilder.GetAggregateSpan(definitionSyntaxChildren)
            ?? CreateSpan(
                state,
                lineNumber,
                definitionStartIndex + 1,
                lineNumber,
                Math.Max(definitionStartIndex + 1, definitionEndExclusive));
        var definition = new DefinitionListDefinition(definitionBlocks);
        definitionNode = new MarkdownSyntaxNode(
            MarkdownSyntaxKind.DefinitionValue,
            definitionSpan,
            RenderDefinitionLiteral(definitionBlocks, definitionLiteral),
            definitionSyntaxChildren,
            associatedObject: definition);
        return definition;
    }

    private sealed class ParsedDefinitionTermLine {
        public ParsedDefinitionTermLine(int zeroBasedLineIndex, int startIndex, int endExclusive, string literal) {
            ZeroBasedLineIndex = zeroBasedLineIndex;
            StartIndex = startIndex;
            EndExclusive = endExclusive;
            Literal = literal ?? string.Empty;
        }

        public int ZeroBasedLineIndex { get; }
        public int StartIndex { get; }
        public int EndExclusive { get; }
        public string Literal { get; }
        public InlineSequence Inlines { get; private set; } = new InlineSequence();
        public MarkdownSourceSpan Span { get; private set; }

        public void Bind(MarkdownReaderOptions options, MarkdownReaderState state) {
            int lineNumber = state.SourceLineOffset + ZeroBasedLineIndex + 1;
            var sourceMap = BuildInlineSourceMapForSingleLine(Literal, lineNumber, StartIndex + 1, state);
            Inlines = ParseInlines(Literal, options, state, sourceMap);
            Span = CreateSpan(
                state,
                lineNumber,
                StartIndex + 1,
                lineNumber,
                Math.Max(StartIndex + 1, EndExclusive));
        }
    }

    private static void ConsumeDefinitionContinuationLines(
        string[] lines,
        ref int index,
        int continuationIndent,
        int absoluteLineOffset,
        List<MarkdownSourceLineSlice> definitionSourceLines,
        bool allowLazyContinuation) {
        if (lines == null || definitionSourceLines == null) {
            return;
        }

        while (index < lines.Length) {
            var line = lines[index] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(line)) {
                int peek = index;
                while (peek < lines.Length && string.IsNullOrWhiteSpace(lines[peek])) {
                    peek++;
                }

                if (peek >= lines.Length || CountLeadingIndentColumns(lines[peek] ?? string.Empty) < continuationIndent) {
                    return;
                }

                while (index < peek) {
                    definitionSourceLines.Add(new MarkdownSourceLineSlice(string.Empty, absoluteLineOffset + index + 1, 1));
                    index++;
                }
                continue;
            }

            if (CountLeadingIndentColumns(line) < continuationIndent) {
                if (!allowLazyContinuation ||
                    !ShouldConsumeMarkdigDefinitionLazyContinuation(lines, index)) {
                    return;
                }

                int firstContentIndex = GetFirstNonWhitespaceIndex(line);
                definitionSourceLines.Add(new MarkdownSourceLineSlice(
                    firstContentIndex < line.Length ? line.Substring(firstContentIndex) : string.Empty,
                    absoluteLineOffset + index + 1,
                    firstContentIndex + 1));
                index++;
                continue;
            }

            var stripped = StripLeadingIndentColumns(line, continuationIndent);
            definitionSourceLines.Add(new MarkdownSourceLineSlice(
                stripped,
                absoluteLineOffset + index + 1,
                GetStartColumnAfterStrippingIndent(line, continuationIndent)));
            index++;
        }
    }

    private static bool ShouldConsumeMarkdigDefinitionLazyContinuation(
        string[] lines,
        int index) {
        if (lines == null || index < 0 || index >= lines.Length) {
            return false;
        }

        var line = lines[index] ?? string.Empty;
        if (string.IsNullOrWhiteSpace(line)) {
            return false;
        }

        if (TryGetMarkdigDefinitionMarker(line, out _, out _)) {
            return false;
        }

        var trimmed = line.TrimStart();
        if (trimmed.Length == 0) {
            return false;
        }

        return true;
    }

    private static int GetFirstNonWhitespaceIndex(string line) {
        if (string.IsNullOrEmpty(line)) {
            return 0;
        }

        int index = 0;
        while (index < line.Length && char.IsWhiteSpace(line[index])) {
            index++;
        }

        return index;
    }

    private static (IReadOnlyList<IMarkdownBlock> Blocks, IReadOnlyList<MarkdownSyntaxNode> SyntaxChildren) ParseDefinitionBody(
        List<MarkdownSourceLineSlice> definitionSourceLines,
        MarkdownReaderOptions options,
        MarkdownReaderState state) {
        if (definitionSourceLines == null || definitionSourceLines.Count == 0) {
            return (Array.Empty<IMarkdownBlock>(), Array.Empty<MarkdownSyntaxNode>());
        }

        var (blocks, syntaxChildren) = ParseNestedMarkdownBlocks(definitionSourceLines, options, state);
        if (blocks.Count > 0) {
            return (blocks, syntaxChildren);
        }

        var paragraphs = ParseParagraphBlocksFromSourceLines(definitionSourceLines, options, state);
        var nodes = new List<MarkdownSyntaxNode>();
        AddParagraphSyntaxNodes(nodes, definitionSourceLines, options, state);
        return (paragraphs, nodes);
    }

    private static int GetStartColumnAfterStrippingIndent(string line, int requiredColumns) {
        if (string.IsNullOrEmpty(line) || requiredColumns <= 0) {
            return 1;
        }

        int rawIndex = 0;
        int consumedColumns = 0;
        while (rawIndex < line.Length && consumedColumns < requiredColumns) {
            if (line[rawIndex] == ' ') {
                consumedColumns++;
                rawIndex++;
                continue;
            }

            if (line[rawIndex] == '\t') {
                consumedColumns = ((consumedColumns / 4) + 1) * 4;
                rawIndex++;
                continue;
            }

            break;
        }

        return rawIndex + 1;
    }

    private static string RenderDefinitionLiteral(IReadOnlyList<IMarkdownBlock> blocks, string fallbackLiteral) {
        if (blocks == null || blocks.Count == 0) {
            return fallbackLiteral ?? string.Empty;
        }

        var sb = new StringBuilder();
        for (int i = 0; i < blocks.Count; i++) {
            var rendered = blocks[i]?.RenderMarkdown();
            if (string.IsNullOrEmpty(rendered)) {
                continue;
            }

            if (sb.Length > 0) {
                sb.Append("\n\n");
            }

            sb.Append(rendered);
        }

        return sb.Length > 0 ? sb.ToString() : fallbackLiteral ?? string.Empty;
    }
}

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
                var markerSpan = CreateSpan(
                    state,
                    lineNumber,
                    idx + 1,
                    lineNumber,
                    idx + 1);
                var definitionSpan = CreateSpan(
                    state,
                    lineNumber,
                    definitionStartIndex + 1,
                    lineNumber,
                    Math.Max(definitionStartIndex + 1, definitionEndExclusive));
                var definitionSourceLines = new List<MarkdownSourceLineSlice> {
                    new MarkdownSourceLineSlice(def, lineNumber, definitionStartIndex + 1)
                };
                var continuationIndentSourceSpans = new List<MarkdownSourceSpan>();
                int next = j + 1;
                ConsumeDefinitionContinuationLines(
                    lines,
                    ref next,
                    CountLeadingIndentColumns(line) + DefinitionContinuationIndentOffset,
                    null,
                    state.SourceLineOffset,
                    definitionSourceLines,
                    continuationIndentSourceSpans,
                    state,
                    options,
                    allowLazyContinuation: false);
                var (definitionBlocks, definitionSyntaxChildren) = ParseDefinitionBody(definitionSourceLines, options, state);
                if (definitionSyntaxChildren.Count > 0) {
                    definitionSpan = MarkdownBlockSyntaxBuilder.GetAggregateSpan(definitionSyntaxChildren) ?? definitionSpan;
                }

                var definitionObject = new DefinitionListDefinition(definitionBlocks);
                if (definitionBlocks.Count == 1 &&
                    definitionBlocks[0] is ParagraphBlock definitionParagraph &&
                    !definitionParagraph.Attributes.IsEmpty) {
                    definitionObject.ForceParagraphHtml = true;
                }

                definitionObject.ReplaceBlankLineSourceSpans(GetDefinitionBlankLineSourceSpans(definitionSourceLines, state));
                definitionObject.ReplaceContinuationIndentSourceSpans(continuationIndentSourceSpans);
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
                var markerNode = new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.DefinitionMarker,
                    markerSpan,
                    ":");
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
                            markerNode,
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

        if (!TryCollectMarkdigDefinitionTerms(lines, i, options, state, out var markerIndex, out var terms)) {
            return false;
        }

        var dl = new DefinitionListBlock();
        dl.SetParsingContext(options, state);
        int cursor = i;
        while (cursor < lines.Length &&
               TryCollectMarkdigDefinitionTerms(lines, cursor, options, state, out markerIndex, out terms)) {
            var definitions = new List<DefinitionListDefinition>();
            var definitionNodes = new List<(MarkdownSyntaxNode Marker, MarkdownSyntaxNode Definition)>();
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
                    out var markerNode,
                    out var definitionNode);

                if (definition != null && markerNode != null && definitionNode != null) {
                    definitions.Add(definition);
                    definitionNodes.Add((markerNode, definitionNode));
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

            var group = new DefinitionListGroup();
            for (int termIndex = 0; termIndex < terms.Count; termIndex++) {
                group.AddTerm(terms[termIndex].TermObject);
            }

            for (int definitionIndex = 0; definitionIndex < definitions.Count; definitionIndex++) {
                group.AddDefinition(definitions[definitionIndex]);
            }

            var groupChildren = new List<MarkdownSyntaxNode>(terms.Count + definitionNodes.Count);
            for (int termIndex = 0; termIndex < terms.Count; termIndex++) {
                var termObject = group.TermItems[termIndex];
                groupChildren.Add(MarkdownBlockSyntaxBuilder.BuildInlineContainerNode(
                    MarkdownSyntaxKind.DefinitionTerm,
                    termObject.Inlines,
                    terms[termIndex].Span,
                    termObject.Markdown,
                    associatedObject: termObject));
            }

            for (int definitionIndex = 0; definitionIndex < definitionNodes.Count; definitionIndex++) {
                groupChildren.Add(definitionNodes[definitionIndex].Marker);
                groupChildren.Add(definitionNodes[definitionIndex].Definition);
            }

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
                TryCollectMarkdigDefinitionTerms(lines, nextGroup, options, state, out _, out _)) {
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
        MarkdownReaderOptions options,
        MarkdownReaderState state,
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

            if (!IsMarkdigDefinitionTermCandidate(line, options)) {
                return false;
            }

            terms.Add(CreateMarkdigDefinitionTerm(line, cursor, options, state));
            cursor++;
        }

        return false;
    }

    private static ParsedDefinitionTermLine CreateMarkdigDefinitionTerm(
        string line,
        int zeroBasedLineIndex,
        MarkdownReaderOptions options,
        MarkdownReaderState state) {
        int termStartIndex = 0;
        while (termStartIndex < line.Length && char.IsWhiteSpace(line[termStartIndex])) {
            termStartIndex++;
        }

        int termEndExclusive = line.Length;
        while (termEndExclusive > termStartIndex && char.IsWhiteSpace(line[termEndExclusive - 1])) {
            termEndExclusive--;
        }

        var sourceLiteral = termEndExclusive > termStartIndex
            ? line.Substring(termStartIndex, termEndExclusive - termStartIndex)
            : string.Empty;
        var literal = sourceLiteral;
        MarkdownAttributeSet attributes = MarkdownAttributeSet.Empty;
        MarkdownSourceSpan? attributeSpan = null;
        string? attributeSourceText = null;
        string consumedWhitespace = string.Empty;

        if (ShouldParseBlockGenericAttributes(options, state)
            && MarkdownGenericAttributeParser.TryConsumeTrailingAttributeBlock(
                sourceLiteral,
                out var literalWithoutAttributeBlock,
                out attributes,
                out var attributeStart,
                out var attributeEnd,
                requireLeadingWhitespace: true)) {
            literal = literalWithoutAttributeBlock;
            if (attributeStart >= literalWithoutAttributeBlock.Length) {
                consumedWhitespace = sourceLiteral.Substring(
                    literalWithoutAttributeBlock.Length,
                    attributeStart - literalWithoutAttributeBlock.Length);
            }

            attributeSourceText = sourceLiteral.Substring(attributeStart, attributeEnd - attributeStart + 1);
            var absoluteAttributeLine = state.SourceLineOffset + zeroBasedLineIndex + 1;
            attributeSpan = CreateSpan(
                state,
                absoluteAttributeLine,
                termStartIndex + attributeStart + 1,
                absoluteAttributeLine,
                termStartIndex + attributeEnd + 1);
        }

        return new ParsedDefinitionTermLine(
            zeroBasedLineIndex,
            termStartIndex,
            termEndExclusive,
            literal,
            attributes,
            attributeSourceText,
            attributeSpan,
            consumedWhitespace);
    }

    private static bool IsMarkdigDefinitionTermCandidate(string line, MarkdownReaderOptions options) {
        if (string.IsNullOrWhiteSpace(line)) return false;
        if (CountLeadingIndentColumns(line) >= 4) return false;
        var trimmed = line.TrimStart();
        if (IsAtxHeading(trimmed, out _, out _)) return false;
        if (IsUnorderedListLine(trimmed, out _, out _, out _)) return false;
        if (IsOrderedListLine(trimmed, options, out _, out _)) return false;
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
        out MarkdownSyntaxNode? markerNode,
        out MarkdownSyntaxNode? definitionNode) {
        var line = lines[index] ?? string.Empty;
        int lineNumber = state.SourceLineOffset + index + 1;
        var markerSpan = CreateSpan(
            state,
            lineNumber,
            markerIndent + 1,
            lineNumber,
            markerIndent + 1);
        markerNode = new MarkdownSyntaxNode(
            MarkdownSyntaxKind.DefinitionMarker,
            markerSpan,
            ":");
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
        var continuationIndentSourceSpans = new List<MarkdownSourceSpan>();
        var consumedLeadingBlankLine = ConsumeDefinitionContinuationLines(
            lines,
            ref next,
            markerIndent + DefinitionContinuationIndentOffset,
            string.IsNullOrEmpty(definitionLiteral) ? markerIndent + 4 : null,
            state.SourceLineOffset,
            definitionSourceLines,
            continuationIndentSourceSpans,
            state,
            options,
            allowLazyContinuation: true);
        index = next;

        if (definitionSourceLines.Count == 0) {
            markerNode = null;
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
        definition.ReplaceBlankLineSourceSpans(GetDefinitionBlankLineSourceSpans(definitionSourceLines, state));
        definition.ReplaceContinuationIndentSourceSpans(continuationIndentSourceSpans);
        if (consumedLeadingBlankLine) {
            definition.ForceParagraphHtml = true;
            definition.HasLeadingBlankLineBeforeBody = true;
        }

        definitionNode = new MarkdownSyntaxNode(
            MarkdownSyntaxKind.DefinitionValue,
            definitionSpan,
            RenderDefinitionLiteral(definitionBlocks, definitionLiteral),
            definitionSyntaxChildren,
            associatedObject: definition);
        return definition;
    }

    private sealed class ParsedDefinitionTermLine {
        public ParsedDefinitionTermLine(
            int zeroBasedLineIndex,
            int startIndex,
            int endExclusive,
            string literal,
            MarkdownAttributeSet attributes,
            string? attributeSourceText,
            MarkdownSourceSpan? attributeSpan,
            string consumedWhitespace) {
            ZeroBasedLineIndex = zeroBasedLineIndex;
            StartIndex = startIndex;
            EndExclusive = endExclusive;
            Literal = literal ?? string.Empty;
            Attributes = attributes ?? MarkdownAttributeSet.Empty;
            AttributeSourceText = attributeSourceText;
            AttributeSpan = attributeSpan;
            ConsumedWhitespace = consumedWhitespace ?? string.Empty;
        }

        public int ZeroBasedLineIndex { get; }
        public int StartIndex { get; }
        public int EndExclusive { get; }
        public string Literal { get; }
        public MarkdownAttributeSet Attributes { get; }
        public string? AttributeSourceText { get; }
        public MarkdownSourceSpan? AttributeSpan { get; }
        public string ConsumedWhitespace { get; }
        public InlineSequence Inlines { get; private set; } = new InlineSequence();
        public MarkdownSourceSpan Span { get; private set; }
        public DefinitionListTerm TermObject { get; private set; } = new DefinitionListTerm();

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
            TermObject = new DefinitionListTerm(Inlines);
            TermObject.SetAttributes(Attributes);
            TermObject.GenericAttributeConsumedWhitespace = ConsumedWhitespace;
            TermObject.SourceSpan = Span;
            MarkdownGenericAttributeSourceSpans.Set(TermObject, AttributeSourceText, AttributeSpan);
        }
    }

    private static bool ConsumeDefinitionContinuationLines(
        string[] lines,
        ref int index,
        int continuationIndent,
        int? firstContinuationIndent,
        int absoluteLineOffset,
        List<MarkdownSourceLineSlice> definitionSourceLines,
        List<MarkdownSourceSpan> continuationIndentSourceSpans,
        MarkdownReaderState state,
        MarkdownReaderOptions options,
        bool allowLazyContinuation) {
        if (lines == null || definitionSourceLines == null) {
            return false;
        }

        bool useFirstContinuationIndent = firstContinuationIndent.HasValue;
        bool hasContent = definitionSourceLines.Any(static sourceLine => !string.IsNullOrWhiteSpace(sourceLine.Text));
        bool consumedLeadingBlankLine = false;
        bool consumedBlankAfterContent = false;
        while (index < lines.Length) {
            var line = lines[index] ?? string.Empty;
            int effectiveContinuationIndent = useFirstContinuationIndent
                ? firstContinuationIndent.GetValueOrDefault()
                : continuationIndent;
            if (allowLazyContinuation && consumedBlankAfterContent) {
                effectiveContinuationIndent = Math.Max(effectiveContinuationIndent, continuationIndent + DefinitionContinuationIndentOffset);
            }

            if (string.IsNullOrWhiteSpace(line)) {
                int peek = index;
                while (peek < lines.Length && string.IsNullOrWhiteSpace(lines[peek])) {
                    peek++;
                }

                int followingContinuationIndent = allowLazyContinuation && hasContent
                    ? Math.Max(effectiveContinuationIndent, continuationIndent + DefinitionContinuationIndentOffset)
                    : effectiveContinuationIndent;
                if (peek >= lines.Length || CountLeadingIndentColumns(lines[peek] ?? string.Empty) < followingContinuationIndent) {
                    return consumedLeadingBlankLine;
                }

                bool blankAfterContent = hasContent;
                if (!hasContent) {
                    consumedLeadingBlankLine = true;
                }

                while (index < peek) {
                    definitionSourceLines.Add(new MarkdownSourceLineSlice(string.Empty, absoluteLineOffset + index + 1, 1));
                    index++;
                }

                consumedBlankAfterContent = blankAfterContent;
                continue;
            }

            if (CountLeadingIndentColumns(line) < effectiveContinuationIndent) {
                if (!allowLazyContinuation ||
                    !ShouldConsumeMarkdigDefinitionLazyContinuation(lines, index, definitionSourceLines, options)) {
                    return consumedLeadingBlankLine;
                }

                int firstContentIndex = GetFirstNonWhitespaceIndex(line);
                definitionSourceLines.Add(new MarkdownSourceLineSlice(
                    firstContentIndex < line.Length ? line.Substring(firstContentIndex) : string.Empty,
                    absoluteLineOffset + index + 1,
                    firstContentIndex + 1,
                    isLazyQuoteContinuation: true));
                hasContent = true;
                index++;
                useFirstContinuationIndent = false;
                consumedBlankAfterContent = false;
                continue;
            }

            int startColumn = GetStartColumnAfterStrippingIndent(line, effectiveContinuationIndent);
            if (startColumn > 1) {
                continuationIndentSourceSpans?.Add(CreateSpan(
                    state,
                    absoluteLineOffset + index + 1,
                    1,
                    absoluteLineOffset + index + 1,
                    startColumn - 1));
            }

            var stripped = StripLeadingIndentColumns(line, effectiveContinuationIndent);
            definitionSourceLines.Add(new MarkdownSourceLineSlice(
                stripped,
                absoluteLineOffset + index + 1,
                startColumn));
            if (!string.IsNullOrWhiteSpace(stripped)) {
                hasContent = true;
            }

            index++;
            useFirstContinuationIndent = false;
            consumedBlankAfterContent = false;
        }

        return consumedLeadingBlankLine;
    }

    private static bool ShouldConsumeMarkdigDefinitionLazyContinuation(
        string[] lines,
        int index,
        List<MarkdownSourceLineSlice> definitionSourceLines,
        MarkdownReaderOptions options) {
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

        if (PreviousDefinitionLineStopsLazyContinuation(definitionSourceLines)) {
            return false;
        }

        if (PreviousDefinitionLinesContainActiveNestedBlock(definitionSourceLines)) {
            if (PreviousDefinitionLinesEndActiveNestedBlockquote(definitionSourceLines) &&
                trimmed.StartsWith(">", StringComparison.Ordinal)) {
                return true;
            }

            if (PreviousDefinitionLinesEndActiveNestedBlockquote(definitionSourceLines) &&
                CurrentLineStartsListBlock(line)) {
                return false;
            }

            if (CurrentLineInterruptsNestedDefinitionBlock(line, options)) {
                return false;
            }
        }

        return true;
    }

    private static bool PreviousDefinitionLineStopsLazyContinuation(List<MarkdownSourceLineSlice> definitionSourceLines) {
        if (definitionSourceLines == null || definitionSourceLines.Count == 0) {
            return false;
        }

        var previous = definitionSourceLines[definitionSourceLines.Count - 1].Text;
        if (TryGetSetextHeadingUnderlineLevel(previous, out _)) {
            return !PreviousDefinitionLinesContainActiveNestedBlock(definitionSourceLines);
        }

        return IsParagraphInterruptingThematicBreakLine(previous) ||
            PreviousDefinitionLinesEndClosedFencedCodeBlock(definitionSourceLines);
    }

    private static bool PreviousDefinitionLinesEndClosedFencedCodeBlock(List<MarkdownSourceLineSlice> definitionSourceLines) {
        bool inFence = false;
        char fenceChar = '\0';
        int fenceLength = 0;
        bool previousNonBlankClosedFence = false;

        for (int i = 0; i < definitionSourceLines.Count; i++) {
            var line = definitionSourceLines[i].Text ?? string.Empty;
            if (string.IsNullOrWhiteSpace(line)) {
                continue;
            }

            previousNonBlankClosedFence = false;
            if (inFence) {
                if (IsCodeFenceClose(line, fenceChar, fenceLength)) {
                    inFence = false;
                    previousNonBlankClosedFence = true;
                }

                continue;
            }

            if (IsCodeFenceOpen(line, out _, out fenceChar, out fenceLength)) {
                inFence = true;
            }
        }

        return previousNonBlankClosedFence && !inFence;
    }

    private static bool PreviousDefinitionLinesContainActiveNestedBlock(List<MarkdownSourceLineSlice> definitionSourceLines) {
        if (definitionSourceLines == null || definitionSourceLines.Count == 0) {
            return false;
        }

        for (int i = definitionSourceLines.Count - 1; i >= 0; i--) {
            var previous = definitionSourceLines[i].Text;
            if (string.IsNullOrWhiteSpace(previous)) {
                return false;
            }

            var trimmed = previous.TrimStart();
            if (trimmed.StartsWith(">", StringComparison.Ordinal) ||
                TryGetUnorderedListMarkerInfo(previous, out _, out _, out _) ||
                TryGetOrderedListMarkerInfo(previous, out _, out _, out _, out _)) {
                return true;
            }
        }

        return false;
    }

    private static bool PreviousDefinitionLinesEndActiveNestedBlockquote(List<MarkdownSourceLineSlice> definitionSourceLines) {
        if (definitionSourceLines == null || definitionSourceLines.Count == 0) {
            return false;
        }

        for (int i = definitionSourceLines.Count - 1; i >= 0; i--) {
            var previous = definitionSourceLines[i].Text;
            if (string.IsNullOrWhiteSpace(previous)) {
                return false;
            }

            if (previous.TrimStart().StartsWith(">", StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
    }

    private static bool CurrentLineInterruptsNestedDefinitionBlock(string line, MarkdownReaderOptions options) {
        if (string.IsNullOrWhiteSpace(line)) {
            return false;
        }

        var trimmed = line.TrimStart();
        return IsAtxHeading(line, out _, out _) ||
            LooksLikeHr(line) ||
            IsCodeFenceOpen(trimmed, out _, out _, out _) ||
            HtmlBlockParser.IsParagraphInterruptingHtmlBlockStart(trimmed, options) ||
            trimmed.StartsWith(">", StringComparison.Ordinal);
    }

    private static bool CurrentLineStartsListBlock(string line) {
        if (string.IsNullOrWhiteSpace(line)) {
            return false;
        }

        var trimmed = line.TrimStart();
        return IsUnorderedListLine(trimmed, out _, out _, out _) ||
            IsOrderedListLine(trimmed, out _, out _);
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

        definitionSourceLines = NormalizeDefinitionFencedCodeSourceLines(definitionSourceLines);

        if (ShouldParseDefinitionBodyAsLiteralParagraph(definitionSourceLines, options)) {
            var literalParagraphs = ParseParagraphBlocksFromSourceLines(definitionSourceLines, options, state);
            var literalNodes = new List<MarkdownSyntaxNode>();
            AddParagraphSyntaxNodes(literalNodes, definitionSourceLines, options, state);
            return (literalParagraphs, literalNodes);
        }

        var definitionBodyState = CloneState(state);
        definitionBodyState.IsMarkdigDefinitionListBody = true;
        RemoveLazyDefinitionContinuationReferenceDefinitions(definitionBodyState, definitionSourceLines);
        var suppressedParagraphAttributeStartLines = GetContinuationParagraphGenericAttributeSuppressionLines(definitionSourceLines);
        var (blocks, syntaxChildren) = ParseNestedMarkdownBlocks(
            definitionSourceLines,
            options,
            definitionBodyState,
            suppressedParagraphAttributeStartLines);
        (blocks, syntaxChildren) = MergeMarkdigDefinitionLazyListContinuations(blocks, syntaxChildren, definitionSourceLines, options, state);
        (blocks, syntaxChildren) = PreserveMarkdigDefinitionLazyParagraphSoftBreaks(blocks, syntaxChildren, definitionSourceLines, options, state);
        if (blocks.Count > 0) {
            return (blocks, syntaxChildren);
        }

        var paragraphs = ParseParagraphBlocksFromSourceLines(definitionSourceLines, options, state);
        var nodes = new List<MarkdownSyntaxNode>();
        AddParagraphSyntaxNodes(nodes, definitionSourceLines, options, state);
        return (paragraphs, nodes);
    }

    private static void RemoveLazyDefinitionContinuationReferenceDefinitions(
        MarkdownReaderState state,
        IReadOnlyList<MarkdownSourceLineSlice> definitionSourceLines) {
        if (state == null ||
            state.LinkRefs.Count == 0 ||
            definitionSourceLines == null ||
            definitionSourceLines.Count == 0) {
            return;
        }

        HashSet<int>? lazyReferenceLines = null;
        for (int i = 0; i < definitionSourceLines.Count; i++) {
            var sourceLine = definitionSourceLines[i];
            if (!sourceLine.IsLazyQuoteContinuation ||
                !StartsWithReferenceDefinitionLikeLabel(sourceLine.Text.TrimStart())) {
                continue;
            }

            lazyReferenceLines ??= new HashSet<int>();
            lazyReferenceLines.Add(sourceLine.AbsoluteLine);
        }

        if (lazyReferenceLines == null || lazyReferenceLines.Count == 0) {
            return;
        }

        foreach (var pair in state.LinkRefs.ToArray()) {
            var definition = pair.Value;
            int? startLine = definition.LabelSourceSpan?.StartLine ?? definition.SourceSpan?.StartLine;
            if (startLine.HasValue && lazyReferenceLines.Contains(startLine.Value)) {
                state.LinkRefs.Remove(pair.Key);
            }
        }
    }

    private static IReadOnlyList<MarkdownSourceSpan> GetDefinitionBlankLineSourceSpans(
        IReadOnlyList<MarkdownSourceLineSlice> definitionSourceLines,
        MarkdownReaderState state) {
        if (definitionSourceLines == null || definitionSourceLines.Count == 0) {
            return Array.Empty<MarkdownSourceSpan>();
        }

        var spans = new List<MarkdownSourceSpan>();
        for (int i = 0; i < definitionSourceLines.Count; i++) {
            var sourceLine = definitionSourceLines[i];
            if (!string.IsNullOrWhiteSpace(sourceLine.Text)) {
                continue;
            }

            spans.Add(CreateSpan(
                state,
                sourceLine.AbsoluteLine,
                sourceLine.StartColumn,
                sourceLine.AbsoluteLine,
                sourceLine.StartColumn));
        }

        return spans;
    }

    private static List<MarkdownSourceLineSlice> NormalizeDefinitionFencedCodeSourceLines(List<MarkdownSourceLineSlice> definitionSourceLines) {
        var normalized = new List<MarkdownSourceLineSlice>(definitionSourceLines.Count);
        bool inFence = false;
        char fenceChar = '\0';
        int fenceLength = 0;
        int fenceColumn = 1;

        for (int i = 0; i < definitionSourceLines.Count; i++) {
            var sourceLine = definitionSourceLines[i];
            if (!inFence) {
                normalized.Add(sourceLine);
                if (IsCodeFenceOpen(sourceLine.Text, out _, out fenceChar, out fenceLength)) {
                    fenceColumn = sourceLine.StartColumn + CountLeadingIndentColumns(sourceLine.Text);
                    inFence = true;
                }

                continue;
            }

            var alignedLine = AlignDefinitionFencedCodeSourceLine(sourceLine, fenceColumn);
            normalized.Add(alignedLine);
            if (IsCodeFenceClose(alignedLine.Text, fenceChar, fenceLength)) {
                inFence = false;
            }
        }

        return normalized;
    }

    private static MarkdownSourceLineSlice AlignDefinitionFencedCodeSourceLine(MarkdownSourceLineSlice sourceLine, int fenceColumn) {
        if (string.IsNullOrEmpty(sourceLine.Text) || sourceLine.StartColumn >= fenceColumn) {
            return sourceLine;
        }

        int columnsToStrip = fenceColumn - sourceLine.StartColumn;
        if (CountLeadingIndentColumns(sourceLine.Text) < columnsToStrip) {
            return sourceLine;
        }

        var stripped = StripLeadingIndentColumns(sourceLine.Text, columnsToStrip);
        int relativeStartColumn = GetStartColumnAfterStrippingIndent(sourceLine.Text, columnsToStrip);
        return new MarkdownSourceLineSlice(
            stripped,
            sourceLine.AbsoluteLine,
            sourceLine.StartColumn + relativeStartColumn - 1,
            sourceLine.IsLazyQuoteContinuation,
            sourceLine.IsQuoteContainerLine);
    }

    private static bool ShouldParseDefinitionBodyAsLiteralParagraph(
        List<MarkdownSourceLineSlice> definitionSourceLines,
        MarkdownReaderOptions options) {
        if (options?.Tables == true || definitionSourceLines == null || definitionSourceLines.Count < 2) {
            return false;
        }

        return LooksLikeTableRow(definitionSourceLines[0].Text)
            && IsAlignmentRow(definitionSourceLines[1].Text);
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

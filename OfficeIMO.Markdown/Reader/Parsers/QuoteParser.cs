namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class QuoteParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            var t = lines[i];
            // Exclude callouts (handled earlier): they start with "> [!"
            if (CountLeadingIndentColumns(t) > 3) return false;
            var quoteMarkerIndex = GetFirstNonWhitespaceIndex(t);
            if (quoteMarkerIndex >= t.Length || t[quoteMarkerIndex] != '>') return false;
            var trimmed = quoteMarkerIndex == 0 ? t : t.Substring(quoteMarkerIndex);
            if (options.Callouts &&
                IsCalloutHeader(trimmed, options, out _, out _)) return false;

            // Collect contiguous quote lines and un-prefix one ">" level
            var inner = new System.Collections.Generic.List<string>();
            var innerSourceLines = new System.Collections.Generic.List<MarkdownSourceLineSlice>();
            var markerSourceSpans = state.CaptureSyntaxTree
                ? new System.Collections.Generic.List<MarkdownSourceSpan>()
                : null;
            int j = i;
            bool sawQuotedLine = false;
            while (j < lines.Length) {
                var ln = lines[j];
                var ltrim = ln.TrimStart();
                if (ltrim.StartsWith(">", StringComparison.Ordinal)) {
                    if (CountLeadingIndentColumns(ln) > 3) break;

                    if (sawQuotedLine
                        && options.Callouts
                        && inner.Count > 0
                        && IsCalloutHeader(ltrim, options, out _, out _)) {
                        break;
                    }

                    // Strip one level
                    var stripped = ltrim.Length >= 2 && ltrim[1] == ' ' ? ltrim.Substring(2) : ltrim.Substring(1);
                    if (inner.Count == 0) {
                        stripped = NormalizeContainerContentIndent(stripped);
                    }
                    if (inner.Count > 0 &&
                        TryNormalizeQuotedNestedQuoteContinuation(inner[inner.Count - 1], stripped, options, out var normalizedNestedQuoteLine)) {
                        stripped = normalizedNestedQuoteLine;
                    } else if (inner.Count > 0 &&
                        TryNormalizeQuotedListContinuationLine(inner[inner.Count - 1], stripped, options, out var normalizedQuotedLine)) {
                        stripped = normalizedQuotedLine;
                    } else if (inner.Count > 0 &&
                        TryNormalizeQuotedIndentedParagraphContinuation(inner[inner.Count - 1], stripped, options, out var normalizedQuotedParagraphLine)) {
                        stripped = normalizedQuotedParagraphLine;
                    }

                    inner.Add(stripped);
                    innerSourceLines.Add(new MarkdownSourceLineSlice(
                        stripped,
                        state.SourceLineOffset + j + 1,
                        GetQuoteContentStartColumn(ln),
                        isQuoteContainerLine: true));
                    var markerSourceSpan = CreateQuoteMarkerSourceSpan(ln, state.SourceLineOffset + j + 1, state);
                    if (markerSourceSpan.HasValue && markerSourceSpans != null) {
                        markerSourceSpans.Add(markerSourceSpan.Value);
                    }
                    sawQuotedLine = true;
                    j++;
                    continue;
                }

                // Lazy continuation: allow a non-quoted line to continue a blockquote paragraph
                // until a blank line followed by a non-quoted line ends the blockquote.
                if (sawQuotedLine) {
                    if (string.IsNullOrWhiteSpace(ln)) {
                        break;
                    }

                    // Only continue lazily when both sides look like paragraph content.
                    // A non-quoted list/item/code starter should end the blockquote instead of being swallowed into it.
                    if (inner.Count > 0) {
                        if (LooksLikeQuoteLazyContinuationPredecessor(inner, inner.Count - 1, options)) {
                            if (!TryNormalizeQuoteLazyContinuationLine(lines, j, options, out var normalizedLazyLine)) break;

                            inner.Add(normalizedLazyLine);
                            innerSourceLines.Add(new MarkdownSourceLineSlice(
                                normalizedLazyLine,
                                state.SourceLineOffset + j + 1,
                                CountLeadingIndentColumns(ln) + 1,
                                isLazyQuoteContinuation: true,
                                isQuoteContainerLine: true));
                            j++;
                            continue;
                        }

                        if (TryNormalizeQuoteLazyContinuationAfterListItem(inner[inner.Count - 1], lines, j, options, out var normalizedListLazyLine)) {
                            inner.Add(normalizedListLazyLine);
                            innerSourceLines.Add(new MarkdownSourceLineSlice(
                                normalizedListLazyLine,
                                state.SourceLineOffset + j + 1,
                                CountLeadingIndentColumns(ln) + 1,
                                isLazyQuoteContinuation: true,
                                isQuoteContainerLine: true));
                            j++;
                            continue;
                        }

                        if (TryNormalizeQuotedNestedQuoteContinuation(inner[inner.Count - 1], ln, options, out var normalizedNestedLazyLine)) {
                            inner.Add(normalizedNestedLazyLine);
                            innerSourceLines.Add(new MarkdownSourceLineSlice(
                                normalizedNestedLazyLine,
                                state.SourceLineOffset + j + 1,
                                CountLeadingIndentColumns(ln) + 1,
                                isLazyQuoteContinuation: true,
                                isQuoteContainerLine: true));
                            j++;
                            continue;
                        }
                    }

                    break;
                }

                break;
            }
            // Recursively parse inner content as a separate document
            var qb = new QuoteBlock();
            IReadOnlyList<MarkdownSyntaxNode> syntaxChildren = Array.Empty<MarkdownSyntaxNode>();
            if (CanParseSemanticQuoteAsSingleParagraph(inner, options, state)) {
                qb.ChildBlocks.Add(new ParagraphBlock(ParseInlines(JoinParagraphLines(inner, options), options, state)));
            } else {
                var parsed = ParseNestedMarkdownBlocks(
                    innerSourceLines,
                    options,
                    state,
                    suppressBlockGenericAttributes: true);
                foreach (var b in parsed.Blocks) qb.ChildBlocks.Add(b);
                syntaxChildren = parsed.SyntaxChildren;
            }
            if (markerSourceSpans != null) {
                qb.ReplaceMarkerSourceSpans(markerSourceSpans);
                qb.SyntaxChildren = syntaxChildren;
            }
            doc.Add(qb); i = j; return true;
        }

        private static bool CanParseSemanticQuoteAsSingleParagraph(
            IReadOnlyList<string> lines,
            MarkdownReaderOptions options,
            MarkdownReaderState state) {
            if (state.CaptureSyntaxTree
                || options.GenericAttributes
                || options.BlockParserExtensions.Count != 0
                || lines == null
                || lines.Count == 0) {
                return false;
            }

            for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
                string line = lines[lineIndex] ?? string.Empty;
                if (string.IsNullOrWhiteSpace(line)
                    || (options.IndentedCodeBlocks && CountLeadingIndentColumns(line) >= 4)
                    || IsQuoteStarter(line)
                    || IsAtxHeading(line, out _, out _)
                    || IsCodeFenceOpen(line, out _, out _, out _)
                    || LooksLikeHr(line)
                    || IsUnorderedListLine(line)
                    || IsOrderedListLine(line, options, out _, out _)
                    || HtmlBlockParser.IsParagraphInterruptingHtmlBlockStart(line, options)
                    || TryGetSetextHeadingUnderlineLevel(line, out _)
                    || LooksLikeReferenceDefinition(line)) {
                    return false;
                }
            }

            return true;
        }

        private static bool LooksLikeReferenceDefinition(string line) {
            int separator = line.IndexOf("]:", StringComparison.Ordinal);
            return separator > 0 && line.LastIndexOf('[', separator) >= 0;
        }
    }

    private static MarkdownSourceSpan? CreateQuoteMarkerSourceSpan(string line, int absoluteLineNumber, MarkdownReaderState state) {
        int markerColumn = GetQuoteMarkerStartColumn(line);
        return CreateSpan(state, absoluteLineNumber, markerColumn, absoluteLineNumber, markerColumn);
    }

    private static int GetQuoteMarkerStartColumn(string line) {
        if (string.IsNullOrEmpty(line)) {
            return 1;
        }

        int column = 1;
        int index = 0;
        while (index < line.Length) {
            char ch = line[index];
            if (ch == ' ') {
                column++;
                index++;
                continue;
            }

            if (ch == '\t') {
                column += 4 - ((column - 1) % 4);
                index++;
                continue;
            }

            break;
        }

        return column;
    }

    private static bool LooksLikeParagraphLine(IReadOnlyList<string>? lines, int index, MarkdownReaderOptions options) {
        if (lines == null || index < 0 || index >= lines.Count) return false;
        var line = lines[index] ?? string.Empty;
        if (string.IsNullOrWhiteSpace(line)) return false;
        if (CountLeadingSpaces(line) >= 4) return false;

        var t = line.TrimStart();

        // Block starters we do not want to lazily continue after.
        if (t.StartsWith(">", StringComparison.Ordinal)) return false;
        if (IsAtxHeading(t, out _, out _)) return false;
        if (LooksLikeHr(t)) return false;
        if (IsCodeFenceOpen(t, out _, out _, out _)) return false;
        if (LooksLikeTableRow(t)) return false;
        if (IsUnorderedListLine(t)) return false;
        if (IsParagraphInterruptingOrderedListLine(t, options)) return false;
        if (ShouldTreatAsDefinitionLine(lines, index, options)) return false;
        if (options.Callouts && IsCalloutHeader("> " + t, options, out _, out _)) return false; // callout marker is quote-prefixed in source

        return true;
    }

    private static bool LooksLikeQuoteLazyContinuationPredecessor(IReadOnlyList<string>? lines, int index, MarkdownReaderOptions options) {
        if (LooksLikeParagraphLine(lines, index, options)) return true;
        if (lines == null || index < 0 || index >= lines.Count) return false;

        var line = lines[index] ?? string.Empty;
        if (string.IsNullOrWhiteSpace(line)) return false;
        if (CountLeadingSpaces(line) >= 4) return false;

        return LooksLikeTableRow(line.TrimStart());
    }

    private static bool TryNormalizeQuotedNestedQuoteContinuation(string? previousLine, string? currentLine, MarkdownReaderOptions options, out string normalized) {
        normalized = currentLine ?? string.Empty;
        if (string.IsNullOrWhiteSpace(previousLine) || string.IsNullOrWhiteSpace(currentLine)) return false;

        int previousQuoteDepth = GetLeadingQuoteMarkerDepth(previousLine!);
        if (previousQuoteDepth == 0) return false;

        int currentQuoteDepth = GetLeadingQuoteMarkerDepth(currentLine);
        if (currentQuoteDepth >= previousQuoteDepth) return false;

        var contentCandidate = StripLeadingQuoteMarkers(currentLine!, currentQuoteDepth);
        if (!TryNormalizeQuoteLazyContinuationLine(new[] { contentCandidate }, 0, options, out var normalizedCurrent)) return false;

        var content = normalizedCurrent.TrimStart();
        if (content.Length == 0) return false;

        var markerPrefix = CreateQuoteMarkerPrefix(previousQuoteDepth);
        normalized = markerPrefix + content;
        return true;
    }

    private static bool TryNormalizeQuoteLazyContinuationLine(IReadOnlyList<string>? lines, int index, MarkdownReaderOptions options, out string normalized) {
        var source = lines != null && index >= 0 && index < lines.Count ? (lines[index] ?? string.Empty) : string.Empty;
        normalized = source;
        if (string.IsNullOrWhiteSpace(source)) return false;

        int leadingSpaces = CountLeadingSpaces(source);
        if (leadingSpaces == 0) {
            return LooksLikeParagraphLine(lines, index, options) ||
                LooksLikeTableRow(source.TrimStart());
        }

        if (leadingSpaces > 4) {
            return false;
        }

        var trimmed = source.TrimStart();
        if (trimmed.Length == 0) return false;
        if (trimmed.StartsWith(">", StringComparison.Ordinal)) return false;
        if (IsAtxHeading(trimmed, out _, out _)) return false;
        if (LooksLikeHr(trimmed)) return false;
        if (IsCodeFenceOpen(trimmed, out _, out _, out _)) return false;
        if (LooksLikeTableRow(trimmed)) return false;
        if (ShouldTreatAsDefinitionLine(lines, index, options)) return false;
        if (options.Callouts && IsCalloutHeader("> " + trimmed, options, out _, out _)) return false;

        if (IsUnorderedListLine(trimmed) || IsParagraphInterruptingOrderedListLine(trimmed, options)) {
            normalized = "\\" + trimmed;
            return true;
        }

        normalized = trimmed;
        return true;
    }

    private static bool TryNormalizeQuoteLazyContinuationAfterListItem(string? previousLine, IReadOnlyList<string>? lines, int index, MarkdownReaderOptions options, out string normalized) {
        normalized = string.Empty;
        if (string.IsNullOrWhiteSpace(previousLine)) return false;
        if (!TryNormalizeQuoteLazyContinuationLine(lines, index, options, out var normalizedLazyLine)) return false;

        var previous = previousLine!;
        if (!IsUnorderedListLine(previous) &&
            !IsOrderedListLine(previous, options, out _, out _, out _, out _)) {
            return false;
        }

        int continuationIndent = GetListContinuationIndent(previous, options);
        normalized = new string(' ', Math.Max(continuationIndent, 1)) + normalizedLazyLine;
        return true;
    }

    private static bool TryNormalizeQuotedListContinuationLine(string? previousLine, string? currentLine, MarkdownReaderOptions options, out string normalized) {
        normalized = currentLine ?? string.Empty;
        if (string.IsNullOrWhiteSpace(previousLine) || string.IsNullOrWhiteSpace(currentLine)) return false;

        var previous = previousLine!;
        if (!IsUnorderedListLine(previous) &&
            !IsOrderedListLine(previous, options, out _, out _, out _, out _)) {
            return false;
        }

        int currentIndent = CountLeadingIndentColumns(currentLine!);
        var trimmed = currentLine!.TrimStart();
        if (trimmed.Length == 0) return false;
        if (trimmed.StartsWith(">", StringComparison.Ordinal)) return false;
        if (IsAtxHeading(trimmed, out _, out _)) return false;
        if (LooksLikeHr(trimmed)) return false;
        if (IsCodeFenceOpen(trimmed, out _, out _, out _)) return false;
        if (LooksLikeTableRow(trimmed)) return false;
        if (ShouldTreatAsDefinitionLine(new[] { currentLine }, 0, options)) return false;
        if (options.Callouts && IsCalloutHeader("> " + trimmed, options, out _, out _)) return false;
        if (IsUnorderedListLine(trimmed) || IsParagraphInterruptingOrderedListLine(trimmed, options)) return false;

        int continuationIndent = GetListContinuationIndent(previous, options);
        if (currentIndent == 0 &&
            TryGetRawListItemContentAfterMarker(previous, out var previousListContent, options) &&
            GetLeadingQuoteMarkerDepth(previousListContent) > 0) {
            normalized = new string(' ', Math.Max(continuationIndent, 1)) + trimmed;
            return true;
        }

        if (currentIndent <= 0) return false;
        if (currentIndent >= continuationIndent) return false;
        if (currentIndent + 1 != continuationIndent) return false;

        normalized = new string(' ', continuationIndent) + trimmed;
        return true;
    }

    private static bool TryNormalizeQuotedIndentedParagraphContinuation(string? previousLine, string? currentLine, MarkdownReaderOptions options, out string normalized) {
        normalized = currentLine ?? string.Empty;
        if (string.IsNullOrWhiteSpace(previousLine) || string.IsNullOrWhiteSpace(currentLine)) return false;

        var previous = previousLine!;
        if (IsUnorderedListLine(previous) ||
            IsOrderedListLine(previous, options, out _, out _, out _, out _)) {
            return false;
        }

        if (!LooksLikeParagraphLine(new[] { previous }, 0, options)) return false;

        int currentIndent = CountLeadingIndentColumns(currentLine!);
        if (currentIndent <= 0 || currentIndent > 4) return false;

        var trimmed = currentLine!.TrimStart();
        if (trimmed.Length == 0) return false;
        if (trimmed.StartsWith(">", StringComparison.Ordinal)) return false;
        if (IsAtxHeading(trimmed, out _, out _)) return false;
        if (LooksLikeHr(trimmed)) return false;
        if (IsCodeFenceOpen(trimmed, out _, out _, out _)) return false;
        if (LooksLikeTableRow(trimmed)) return false;
        if (ShouldTreatAsDefinitionLine(new[] { currentLine }, 0, options)) return false;
        if (options.Callouts && IsCalloutHeader("> " + trimmed, options, out _, out _)) return false;
        if (IsUnorderedListLine(trimmed) || IsParagraphInterruptingOrderedListLine(trimmed, options)) return false;

        normalized = trimmed;
        return true;
    }

    private static int GetLeadingQuoteMarkerDepth(string? line) {
        if (string.IsNullOrWhiteSpace(line)) return 0;

        int depth = 0;
        int index = 0;
        while (index < line!.Length) {
            while (index < line.Length && line[index] == ' ') index++;
            if (index >= line.Length || line[index] != '>') break;
            depth++;
            index++;
            if (index < line.Length && line[index] == ' ') index++;
        }

        return depth;
    }

    private static string StripLeadingQuoteMarkers(string line, int markerDepth) {
        if (string.IsNullOrEmpty(line) || markerDepth <= 0) return line ?? string.Empty;

        int index = 0;
        int stripped = 0;
        while (index < line.Length && stripped < markerDepth) {
            while (index < line.Length && line[index] == ' ') index++;
            if (index >= line.Length || line[index] != '>') break;
            stripped++;
            index++;
            if (index < line.Length && line[index] == ' ') index++;
        }

        return index >= line.Length ? string.Empty : line.Substring(index);
    }

    private static string CreateQuoteMarkerPrefix(int markerDepth) {
        if (markerDepth <= 0) return string.Empty;
        return new string('>', markerDepth) + " ";
    }
}

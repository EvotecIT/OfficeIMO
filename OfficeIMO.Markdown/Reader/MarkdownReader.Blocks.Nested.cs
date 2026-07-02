using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static bool IsParagraphInterruptingOrderedListLine(string line, MarkdownReaderOptions? options = null) {
        if (!IsOrderedListLine(line, options, out _, out int number, out string content, out _)) return false;
        return number == 1 && !string.IsNullOrWhiteSpace(content);
    }

    private static bool IsParagraphInterruptingUnorderedListLine(string line) {
        return IsUnorderedListLine(line, out _, out _, out string content)
               && !string.IsNullOrWhiteSpace(content);
    }

    private static bool LastCollectedLinePreservesIndentedContinuation(List<string> collected, MarkdownReaderOptions? options = null) {
        if (collected == null || collected.Count == 0) return false;

        for (int i = collected.Count - 1; i >= 0; i--) {
            var line = collected[i];
            if (string.IsNullOrWhiteSpace(line)) continue;
            if (!IsOrderedListLine(line, options, out _, out int number, out _, out _)) return false;
            return number != 1;
        }

        return false;
    }

    private static List<string> ConsumeListContinuationLines(
        string[] lines,
        ref int nextIndex,
        int continuationIndent,
        string initialContent,
        MarkdownReaderOptions options,
        bool breakOnAnyOrderedListLine = false,
        List<MarkdownSourceLineSlice>? sourceLines = null,
        int absoluteLineOffset = 0,
        int initialLineIndex = -1,
        int initialStartColumn = 1,
        MarkdownReaderState? state = null) {
        if (lines == null) return new List<string> { initialContent ?? string.Empty };
        if (nextIndex < 0) nextIndex = 0;

        var collected = new List<string> { initialContent ?? string.Empty };
        if (sourceLines != null) {
            int initialAbsoluteLine = initialLineIndex >= 0
                ? absoluteLineOffset + initialLineIndex + 1
                : absoluteLineOffset + nextIndex + 1;
            sourceLines.Add(new MarkdownSourceLineSlice(initialContent ?? string.Empty, initialAbsoluteLine, initialStartColumn));
        }

        int k = nextIndex;

        while (k < lines.Length) {
            var line = lines[k] ?? string.Empty;
            bool collectingLeadFencedCode = TryGetOpenLeadFencedCode(collected, out _, out _, out _);

            if (collectingLeadFencedCode) {
                if (string.IsNullOrWhiteSpace(line)) {
                    collected.Add(string.Empty);
                    sourceLines?.Add(new MarkdownSourceLineSlice(string.Empty, absoluteLineOffset + k + 1, 1));
                    k++;
                    continue;
                }

                int fencedIndentColumns = CountLeadingIndentColumns(line);
                if (fencedIndentColumns < continuationIndent) {
                    break;
                }

                string fencedContent = StripLeadingIndentColumns(line, continuationIndent);
                int fencedStartColumn = continuationIndent + 1;
                sourceLines?.Add(new MarkdownSourceLineSlice(fencedContent, absoluteLineOffset + k + 1, fencedStartColumn));
                collected.Add(fencedContent);
                k++;
                continue;
            }

            int lineIndentColumns = CountLeadingIndentColumns(line);
            bool underIndentedMarkerContinuation = options.StrictListIndentation
                && lineIndentColumns > 3
                && lineIndentColumns < continuationIndent;
            bool breakOnOrderedListLine = breakOnAnyOrderedListLine ||
                IsMarkdigDefinitionLazyOrderedListBoundary(state, k, line, options);

            // Stop before the next list item (including nested items).
            if (!underIndentedMarkerContinuation &&
                (IsUnorderedListLine(line, out _, out _, out _, out _) ||
                (breakOnOrderedListLine ? IsOrderedListLine(line, options, out _, out _, out _, out _) : IsParagraphInterruptingOrderedListLine(line, options)))) {
                break;
            }

            // Stop before nested blocks; they are handled as child blocks of the list item.
            if (CountLeadingIndentColumns(line) >= continuationIndent) {
                var slice = StripLeadingIndentColumns(line, continuationIndent);
                var sliceTrim = slice.TrimStart();
                if (state != null
                    && TryConsumeNestedStandaloneGenericAttributeLine(
                        lines,
                        k,
                        continuationIndent,
                        options,
                        state,
                        NestedStandaloneGenericAttributeTarget.Any,
                        out _,
                        out _,
                        out _)) {
                    break;
                }

                if (IsCodeFenceOpen(slice, out _, out _, out _)) break;
                if (IsCustomContainerOpeningLine(slice, options)) break;
                if (sliceTrim.StartsWith(">")) break;

                if (options.HtmlBlocks && sliceTrim.StartsWith("<")) {
                    // Avoid breaking on angle-bracket autolinks like "<https://...>".
                    if (!TryParseAngleAutolink(sliceTrim, 0, out _, out _, out _)) break;
                }

                // Indented code block inside list item: continuationIndent + 4 spaces.
                if (options.IndentedCodeBlocks) {
                    if (lineIndentColumns >= continuationIndent + 4 && !LastCollectedLinePreservesIndentedContinuation(collected, options)) break;
                }

                // Table inside list item: a pipe row followed by an alignment/row.
                if (options.Tables && LooksLikeTableRow(sliceTrim)) {
                    int peek = k + 1;
                    if (peek < lines.Length && CountLeadingIndentColumns(lines[peek] ?? string.Empty) >= continuationIndent) {
                        var nextSlice = StripLeadingIndentColumns(lines[peek] ?? string.Empty, continuationIndent).TrimStart();
                        // Reduce false positives: require an alignment row, or explicit outer pipes on both rows.
                        bool curOuter = sliceTrim.Length > 0 && sliceTrim[0] == '|' && sliceTrim[sliceTrim.Length - 1] == '|';
                        bool nextOuter = nextSlice.Length > 0 && nextSlice[0] == '|' && nextSlice[nextSlice.Length - 1] == '|';
                        if (IsAlignmentRow(nextSlice) || (curOuter && nextOuter)) break;
                    }
                }
            }

            if (string.IsNullOrWhiteSpace(line)) {
                if (collected.Count == 1 && string.IsNullOrWhiteSpace(collected[0])) {
                    break;
                }

                // Keep blank lines only if followed by an indented continuation line; otherwise end item.
                int peek = k + 1;
                if (peek >= lines.Length) break;
                var next = lines[peek] ?? string.Empty;
                bool nextBreakOnOrderedListLine = breakOnAnyOrderedListLine ||
                    IsMarkdigDefinitionLazyOrderedListBoundary(state, peek, next, options);
                if (IsUnorderedListLine(next, out _, out _, out _, out _) ||
                    (nextBreakOnOrderedListLine ? IsOrderedListLine(next, options, out _, out _, out _, out _) : IsParagraphInterruptingOrderedListLine(next, options))) {
                    break;
                }
                int nextIndentColumns = CountLeadingIndentColumns(next);
                if (nextIndentColumns < continuationIndent) break;

                collected.Add(string.Empty);
                sourceLines?.Add(new MarkdownSourceLineSlice(string.Empty, absoluteLineOffset + k + 1, 1));
                k++;
                continue;
            }

            int indentColumns = lineIndentColumns;
            if (indentColumns < continuationIndent) {
                if (collected.Count > 0 &&
                    !string.IsNullOrWhiteSpace(collected[collected.Count - 1]) &&
                    LooksLikeParagraphLine(collected, collected.Count - 1, options) &&
                    TryNormalizeListLazyContinuationLine(lines, k, options, breakOnOrderedListLine, out var normalizedLazyLine)) {
                    collected.Add(normalizedLazyLine);
                    sourceLines?.Add(new MarkdownSourceLineSlice(
                        normalizedLazyLine,
                        absoluteLineOffset + k + 1,
                        indentColumns + 1,
                        isLazyQuoteContinuation: true));
                    k++;
                    continue;
                }

                break;
            }

            // Strip the required indent; keep the remainder as-is (including additional indentation).
            string cont = StripLeadingIndentColumns(line, continuationIndent);
            int startColumn = continuationIndent + 1;
            startColumn += CountLeadingIndentColumns(cont);
            cont = cont.TrimStart();
            collected.Add(cont);
            sourceLines?.Add(new MarkdownSourceLineSlice(cont, absoluteLineOffset + k + 1, startColumn));
            k++;
        }

        nextIndex = k;
        return collected;
    }

    private static bool IsMarkdigDefinitionLazyOrderedListBoundary(
        MarkdownReaderState? state,
        int lineIndex,
        string line,
        MarkdownReaderOptions? options = null) {
        return state?.IsMarkdigDefinitionListBody == true &&
            state.LazyQuoteContinuationLines.Contains(lineIndex) &&
            IsOrderedListLine(line, options, out _, out _, out _, out _);
    }

    private static bool TryGetOpenLeadFencedCode(
        IReadOnlyList<string>? collected,
        out string language,
        out char fenceChar,
        out int fenceLength) {
        language = string.Empty;
        fenceChar = '\0';
        fenceLength = 0;

        if (collected == null || collected.Count == 0) {
            return false;
        }

        if (!IsCodeFenceOpen(collected[0] ?? string.Empty, out language, out fenceChar, out fenceLength)) {
            return false;
        }

        for (int i = 1; i < collected.Count; i++) {
            if (IsCodeFenceClose(collected[i] ?? string.Empty, fenceChar, fenceLength)) {
                return false;
            }
        }

        return true;
    }

    private static bool TryNormalizeListLazyContinuationLine(IReadOnlyList<string>? lines, int index, MarkdownReaderOptions options, bool breakOnAnyOrderedListLine, out string normalized) {
        var source = lines != null && index >= 0 && index < lines.Count ? (lines[index] ?? string.Empty) : string.Empty;
        normalized = source;
        if (string.IsNullOrWhiteSpace(source)) return false;

        var trimmed = source.TrimStart();
        if (trimmed.Length == 0) return false;
        if (trimmed.StartsWith(">")) return false;
        if (IsAtxHeading(trimmed, out _, out _)) return false;
        if (LooksLikeHr(trimmed)) return false;
        if (IsCodeFenceOpen(trimmed, out _, out _, out _)) return false;
        if (LooksLikeTableRow(trimmed)) return false;
        if (ShouldTreatAsDefinitionLine(lines, index, options)) return false;
        if (options.Callouts && IsCalloutHeader("> " + trimmed, options, out _, out _)) return false;
        int sourceIndentColumns = CountLeadingIndentColumns(source);
        if (IsUnorderedListLine(trimmed, out _, out _, out _, out _)) {
            if (sourceIndentColumns <= 3) return false;
            normalized = trimmed;
            return true;
        }

        if (breakOnAnyOrderedListLine ? IsOrderedListLine(trimmed, options, out _, out _, out _, out _) : IsParagraphInterruptingOrderedListLine(trimmed, options)) {
            if (sourceIndentColumns <= 3) return false;
            normalized = trimmed;
            return true;
        }

        if (options.HtmlBlocks && trimmed.StartsWith("<") && !TryParseAngleAutolink(trimmed, 0, out _, out _, out _)) {
            return false;
        }

        normalized = trimmed;
        return true;
    }

    private static bool TryParseNestedFencedCodeBlock(
        string[] lines,
        ref int index,
        int continuationIndent,
        MarkdownReaderOptions options,
        MarkdownReaderState state,
        out IMarkdownBlock? block) {
        block = null;
        if (lines == null || index < 0 || index >= lines.Length) return false;
        if (!options.FencedCode) return false;

        var attributeSourceText = string.Empty;
        MarkdownAttributeSet attributeSet = MarkdownAttributeSet.Empty;
        MarkdownSourceSpan? attributeSourceSpan = null;
        if (TryConsumeNestedStandaloneGenericAttributeLine(
                lines,
                index,
                continuationIndent,
                options,
                state,
                NestedStandaloneGenericAttributeTarget.FencedCode,
                out attributeSet,
                out attributeSourceText,
                out attributeSourceSpan)) {
            index++;
            if (index >= lines.Length) {
                return false;
            }
        }

        string line = lines[index] ?? string.Empty;
        int indent = CountLeadingIndentColumns(line);
        if (indent < continuationIndent) return false;

        string first = StripLeadingIndentColumns(line, continuationIndent);
        int openingFenceIndent = CountLeadingIndentColumns(first);
        if (openingFenceIndent > 3) return false;

        if (!IsCodeFenceOpen(first, out string language, out char fenceChar, out int fenceLen, out int fenceIndentColumns, out int infoPaddingColumns, out int infoPaddingCharacters)) return false;

        int j = index + 1;
        var code = new StringBuilder();
        bool hasClosingFence = false;
        int closingFenceIndentColumns = 0;
        int closingFenceLength = fenceLen;
        while (j < lines.Length) {
            string raw = lines[j] ?? string.Empty;
            int ind = CountLeadingIndentColumns(raw);
            string sliced = ind >= continuationIndent ? StripLeadingIndentColumns(raw, continuationIndent) : raw.TrimStart();
            if (TryGetCodeFenceCloseInfo(sliced, fenceChar, fenceLen, out closingFenceIndentColumns, out closingFenceLength)) {
                hasClosingFence = true;
                j++;
                break;
            }
            int contentIndentToStrip = Math.Min(openingFenceIndent, CountLeadingIndentColumns(sliced));
            if (contentIndentToStrip > 0) {
                sliced = StripLeadingIndentColumns(sliced, contentIndentToStrip);
            }
            code.AppendLine(sliced);
            j++;
        }

        var contentLineCount = Math.Max(0, j - index - (hasClosingFence ? 2 : 1));
        var content = RemoveSingleTrailingLineEnding(code.ToString());
        string? caption = null;
        // Optional caption line (indented like other nested content)
        if (j < lines.Length) {
            var capLine = lines[j] ?? string.Empty;
            if (CountLeadingIndentColumns(capLine) >= continuationIndent) {
                var capSlice = StripLeadingIndentColumns(capLine, continuationIndent);
                if (TryParseCaption(capSlice, out var cap)) { caption = cap; j++; }
            }
        }

        block = CreateParsedFencedBlock(
            language,
            content,
            isFenced: true,
            caption,
            options,
            fenceIndentColumns,
            fenceLen,
            infoPaddingColumns,
            infoPaddingCharacters,
            fenceChar,
            hasClosingFence,
            openingFenceIndent + closingFenceIndentColumns,
            closingFenceLength,
            contentLineCount);
        if (!attributeSet.IsEmpty && block is MarkdownObject markdownObject && markdownObject.Attributes.IsEmpty) {
            markdownObject.SetAttributes(attributeSet);
            MarkdownGenericAttributeSourceSpans.Set(markdownObject, attributeSourceText, attributeSourceSpan);
        }

        index = j;
        return true;
    }

    private enum NestedStandaloneGenericAttributeTarget {
        Any,
        FencedCode,
        List,
        Paragraph,
        CustomContainer,
        Table
    }

    private static bool TryConsumeNestedStandaloneGenericAttributeLine(
        string[] lines,
        int index,
        int continuationIndent,
        MarkdownReaderOptions options,
        MarkdownReaderState state,
        NestedStandaloneGenericAttributeTarget target,
        out MarkdownAttributeSet attributes,
        out string sourceText,
        out MarkdownSourceSpan? sourceSpan) {
        attributes = MarkdownAttributeSet.Empty;
        sourceText = string.Empty;
        sourceSpan = null;

        if (!ShouldParseNestedStandaloneGenericAttributes(options, state, index)
            || lines == null
            || index < 0
            || index + 1 >= lines.Length) {
            return false;
        }

        var line = lines[index] ?? string.Empty;
        if (CountLeadingIndentColumns(line) < continuationIndent) {
            return false;
        }

        var slice = StripLeadingIndentColumns(line, continuationIndent);
        var attributeIndent = CountLeadingIndentColumns(slice);
        if (attributeIndent > 3) {
            return false;
        }

        var attributeCandidate = StripLeadingIndentColumns(slice, attributeIndent).TrimEnd();
        if (!MarkdownGenericAttributeParser.TryConsumeLeadingAttributeBlock(
                attributeCandidate,
                out var remaining,
                out attributes,
                out var consumedLength)
            || !string.IsNullOrWhiteSpace(remaining)
            || consumedLength != attributeCandidate.Length
            || attributes.IsEmpty) {
            attributes = MarkdownAttributeSet.Empty;
            return false;
        }

        if (!IsNestedStandaloneGenericAttributeTarget(lines, index + 1, continuationIndent, options, state, target)) {
            attributes = MarkdownAttributeSet.Empty;
            return false;
        }

        sourceText = attributeCandidate.Substring(0, consumedLength);
        var absoluteLine = state.SourceLineOffset + index + 1;
        var startColumn = continuationIndent + attributeIndent + 1;
        sourceSpan = CreateSpan(
            state,
            absoluteLine,
            startColumn,
            absoluteLine,
            startColumn + sourceText.Length - 1);
        return true;
    }

    private static bool IsNestedStandaloneGenericAttributeTarget(
        string[] lines,
        int index,
        int continuationIndent,
        MarkdownReaderOptions options,
        MarkdownReaderState state,
        NestedStandaloneGenericAttributeTarget target) {
        if (lines == null || index < 0 || index >= lines.Length) {
            return false;
        }

        var next = lines[index] ?? string.Empty;
        if (CountLeadingIndentColumns(next) < continuationIndent) {
            return false;
        }

        var nextSlice = StripLeadingIndentColumns(next, continuationIndent);
        if ((target == NestedStandaloneGenericAttributeTarget.Any || target == NestedStandaloneGenericAttributeTarget.FencedCode)
            && IsCodeFenceOpen(nextSlice, out _, out _, out _)) {
            return true;
        }

        if ((target == NestedStandaloneGenericAttributeTarget.Any || target == NestedStandaloneGenericAttributeTarget.List)
            && ((options.OrderedLists && IsOrderedListLine(nextSlice, options, out _, out _))
                || (options.UnorderedLists && IsUnorderedListLine(nextSlice, out _, out _, out _, out _)))) {
            return true;
        }

        if ((target == NestedStandaloneGenericAttributeTarget.Any || target == NestedStandaloneGenericAttributeTarget.Paragraph)
            && options.Paragraphs
            && LooksLikeParagraphLine(lines, index, options)) {
            return true;
        }

        if ((target == NestedStandaloneGenericAttributeTarget.Any || target == NestedStandaloneGenericAttributeTarget.CustomContainer)
            && options.CustomContainers
            && IsCustomContainerOpeningLine(nextSlice, options)) {
            return true;
        }

        if (target != NestedStandaloneGenericAttributeTarget.Any && target != NestedStandaloneGenericAttributeTarget.Table) {
            return false;
        }

        var tableLines = new List<string>();
        for (int i = index; i < lines.Length; i++) {
            var line = lines[i] ?? string.Empty;
            if (CountLeadingIndentColumns(line) < continuationIndent) {
                break;
            }

            var sliced = StripLeadingIndentColumns(line, continuationIndent);
            if (string.IsNullOrWhiteSpace(sliced)) {
                break;
            }

            tableLines.Add(sliced);
            if (tableLines.Count >= 2 && !LooksLikeTableRow(sliced.TrimStart()) && !IsAlignmentRow(sliced.TrimStart())) {
                break;
            }
        }

        return tableLines.Count > 0 && StartsTable(tableLines.ToArray(), 0, options, state);
    }

    private static bool TryParseNestedIndentedCodeBlock(string[] lines, ref int index, int continuationIndent, MarkdownReaderOptions options, out CodeBlock? block) {
        block = null;
        if (lines == null || index < 0 || index >= lines.Length) return false;
        if (!options.IndentedCodeBlocks) return false;

        string line = lines[index] ?? string.Empty;
        if (string.IsNullOrWhiteSpace(line)) return false;

        int spaces = CountLeadingIndentColumns(line);
        int required = continuationIndent + 4;
        if (spaces < required) return false;

        int j = index;
        var sb = new StringBuilder();
        while (j < lines.Length) {
            string cur = lines[j] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(cur)) {
                if (!HasIndentedCodeContinuationAfterBlankLines(lines, j, required)) break;
                sb.AppendLine();
                j++;
                continue;
            }

            int curSpaces = CountLeadingIndentColumns(cur);
            if (curSpaces < required) break;
            sb.AppendLine(StripLeadingIndentColumns(cur, required));
            j++;
        }

        block = new CodeBlock(string.Empty, RemoveSingleTrailingLineEnding(sb.ToString()), isFenced: false);
        index = j;
        return true;
    }

    private static bool TryParseNestedQuoteBlock(string[] lines, ref int index, int itemLevelAbs, int continuationIndent, MarkdownReaderOptions options, MarkdownReaderState state, out QuoteBlock? quote) {
        quote = null;
        if (lines == null || index < 0 || index >= lines.Length) return false;

        string line = lines[index] ?? string.Empty;
        if (CountLeadingIndentColumns(line) < continuationIndent) return false;
        string slice = StripLeadingIndentColumns(line, continuationIndent);
        if (CountLeadingIndentColumns(slice) > 3) return false;
        if (!slice.TrimStart().StartsWith(">")) return false;

        int j = index;
        var collected = new List<string>();
        var collectedSourceLines = new List<MarkdownSourceLineSlice>();
        bool sawQuotedLine = false;
        string? lastQuoteContent = null;
        while (j < lines.Length) {
            string raw = lines[j] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(raw)) {
                break;
            }

            if (CountLeadingIndentColumns(raw) < continuationIndent) break;
            string part = StripLeadingIndentColumns(raw, continuationIndent);

            if (string.IsNullOrWhiteSpace(part)) {
                break;
            }

            if (CountLeadingIndentColumns(part) <= 3 && part.TrimStart().StartsWith(">")) {
                int markerStartColumn = continuationIndent + CountLeadingIndentColumns(part) + 1;
                string quoteContent = StripSingleQuoteMarker(part);
                if (TryNormalizeQuotedListContinuationLine(lastQuoteContent, quoteContent, options, out var normalizedQuotedLine)) {
                    quoteContent = normalizedQuotedLine;
                } else if (TryNormalizeQuotedIndentedParagraphContinuation(lastQuoteContent, quoteContent, options, out var normalizedQuotedParagraphLine)) {
                    quoteContent = normalizedQuotedParagraphLine;
                }

                string collectedLine = "> " + quoteContent;
                collected.Add(collectedLine);
                collectedSourceLines.Add(new MarkdownSourceLineSlice(
                    collectedLine,
                    state.SourceLineOffset + j + 1,
                    markerStartColumn));
                sawQuotedLine = true;
                lastQuoteContent = quoteContent;
                j++;
                continue;
            }

            // Match the top-level quote parser's lazy continuation behavior inside list items too.
            if (!sawQuotedLine) break;
            if (IsNestedListLineForParentItem(raw, itemLevelAbs, continuationIndent, options)) break;
            var previousQuoteContent = lastQuoteContent;
            if (previousQuoteContent == null || previousQuoteContent.Length == 0) break;
            var quoteContext = new[] { previousQuoteContent, part };
            if (!LooksLikeParagraphLine(quoteContext, 0, options) ||
                !TryNormalizeQuoteLazyContinuationLine(quoteContext, 1, options, out var normalizedLazyLine)) break;

            collected.Add(normalizedLazyLine);
            collectedSourceLines.Add(new MarkdownSourceLineSlice(
                normalizedLazyLine,
                state.SourceLineOffset + j + 1,
                CountLeadingIndentColumns(raw) + 1));
            lastQuoteContent = normalizedLazyLine;
            j++;
        }

        if (collected.Count == 0) return false;

        var nestedQuoteOptions = GetNestedListQuoteOptions(options);
        var (blocks, _) = ParseNestedMarkdownBlocks(collectedSourceLines, nestedQuoteOptions, state);
        if (blocks.Count > 0 && blocks[0] is QuoteBlock parsedQuote) {
            quote = parsedQuote;
            index = j;
            return true;
        }
        return false;
    }

    private static MarkdownReaderOptions GetNestedListQuoteOptions(MarkdownReaderOptions options) {
        if (!options.Callouts
            || options.CalloutTitleMode != MarkdownCalloutTitleMode.MarkdigCompatible) {
            return options;
        }

        var nestedOptions = CloneOptionsWithoutFrontMatter(options);
        nestedOptions.Callouts = false;
        return nestedOptions;
    }

    private static bool IsNestedListLineForParentItem(string line, int itemLevelAbs, int continuationIndent, MarkdownReaderOptions options) {
        if (string.IsNullOrWhiteSpace(line)) return false;
        if (CountLeadingIndentColumns(line) < continuationIndent) return false;

        if (options.OrderedLists &&
            IsOrderedListLine(line, options, out int orderedLevelAbs, out _, out _, out _) &&
            orderedLevelAbs >= itemLevelAbs + 1) {
            return true;
        }

        if (options.UnorderedLists &&
            IsUnorderedListLine(line, out int unorderedLevelAbs, out _, out _, out _) &&
            unorderedLevelAbs >= itemLevelAbs + 1) {
            return true;
        }

        return false;
    }

    private static string StripSingleQuoteMarker(string line) {
        if (string.IsNullOrEmpty(line)) return string.Empty;
        var trimmed = line.TrimStart();
        if (!trimmed.StartsWith(">")) return trimmed;
        return trimmed.Length >= 2 && trimmed[1] == ' ' ? trimmed.Substring(2) : trimmed.Substring(1);
    }

    private static int GetQuoteContentStartColumn(string line) {
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

        if (index < line.Length && line[index] == '>') {
            column++;
            index++;
        }

        if (index < line.Length && line[index] == ' ') {
            column++;
        }

        return column;
    }

    private static bool TryParseNestedCustomContainerBlock(
        string[] lines,
        ref int index,
        int continuationIndent,
        MarkdownReaderOptions options,
        MarkdownReaderState state,
        out CustomContainerBlock? container,
        out MarkdownSyntaxNode? syntaxNode) {

        container = null;
        syntaxNode = null;
        if (lines == null || index < 0 || index >= lines.Length || !options.CustomContainers) return false;

        var attributeSourceText = string.Empty;
        MarkdownAttributeSet attributeSet = MarkdownAttributeSet.Empty;
        MarkdownSourceSpan? attributeSourceSpan = null;
        if (TryConsumeNestedStandaloneGenericAttributeLine(
                lines,
                index,
                continuationIndent,
                options,
                state,
                NestedStandaloneGenericAttributeTarget.CustomContainer,
                out attributeSet,
                out attributeSourceText,
                out attributeSourceSpan)) {
            index++;
            if (index >= lines.Length) {
                return false;
            }
        }

        string line = lines[index] ?? string.Empty;
        if (CountLeadingIndentColumns(line) < continuationIndent) return false;

        string slice = StripLeadingIndentColumns(line, continuationIndent);
        if (!IsCustomContainerOpeningLine(slice, options)) return false;

        var collected = new List<string>();
        int j = index;
        while (j < lines.Length) {
            string raw = lines[j] ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(raw) && CountLeadingIndentColumns(raw) < continuationIndent) {
                break;
            }

            collected.Add(string.IsNullOrWhiteSpace(raw)
                ? string.Empty
                : StripLeadingIndentColumns(raw, continuationIndent));
            j++;
        }

        if (!CustomContainerParser.TryGetContainerLineCount(collected, 0, out var lineCount) || lineCount <= 0) {
            return false;
        }

        var endIndex = Math.Min(lines.Length, index + lineCount);
        var sourceLines = BuildListItemNestedSourceLines(lines, continuationIndent, index, endIndex, state);
        var (blocks, syntaxChildren) = ParseNestedMarkdownBlocks(sourceLines, options, state);
        if (blocks.Count == 0 || blocks[0] is not CustomContainerBlock parsedContainer) {
            return false;
        }

        container = parsedContainer;
        if (!attributeSet.IsEmpty && container.Attributes.IsEmpty) {
            container.SetAttributes(attributeSet);
            MarkdownGenericAttributeSourceSpans.Set(container, attributeSourceText, attributeSourceSpan);
        }

        if (syntaxChildren.Count > 0 && syntaxChildren[0].Kind == MarkdownSyntaxKind.CustomContainer) {
            syntaxNode = syntaxChildren[0];
        }

        index = endIndex;
        return true;
    }

    private static bool TryParseNestedTableBlock(string[] lines, ref int index, int continuationIndent, MarkdownReaderOptions options, MarkdownReaderState state, out TableBlock? table) {
        table = null;
        if (lines == null || index < 0 || index >= lines.Length) return false;
        if (!options.Tables) return false;

        var attributeSourceText = string.Empty;
        MarkdownAttributeSet attributeSet = MarkdownAttributeSet.Empty;
        MarkdownSourceSpan? attributeSourceSpan = null;
        if (TryConsumeNestedStandaloneGenericAttributeLine(
                lines,
                index,
                continuationIndent,
                options,
                state,
                NestedStandaloneGenericAttributeTarget.Table,
                out attributeSet,
                out attributeSourceText,
                out attributeSourceSpan)) {
            index++;
            if (index >= lines.Length) {
                return false;
            }
        }

        string line = lines[index] ?? string.Empty;
        if (CountLeadingIndentColumns(line) < continuationIndent) return false;
        string slice = StripLeadingIndentColumns(line, continuationIndent);
        if (!LooksLikeTableRow(slice.TrimStart())) return false;

        int j = index;
        var collected = new List<string>();
        var collectedSourceLines = new List<MarkdownSourceLineSlice>();
        while (j < lines.Length) {
            string raw = lines[j] ?? string.Empty;
            if (CountLeadingIndentColumns(raw) < continuationIndent) break;
            string part = StripLeadingIndentColumns(raw, continuationIndent);
            if (string.IsNullOrWhiteSpace(part)) break;
            // Stop when the row no longer looks table-ish.
            if (!LooksLikeTableRow(part.TrimStart()) && !IsAlignmentRow(part.TrimStart())) break;
            collected.Add(part);
            collectedSourceLines.Add(new MarkdownSourceLineSlice(
                part,
                state.SourceLineOffset + j + 1,
                continuationIndent + 1));
            j++;
        }

        if (collected.Count == 0) return false;
        var (blocks, _) = ParseNestedMarkdownBlocks(collectedSourceLines, options, state);
        if (blocks.Count == 0 || blocks[0] is not TableBlock parsedTable) {
            if (!TryParseCollectedNestedBlock(collected, options, state, index, out TableBlock? fallbackTable) || fallbackTable == null) {
                return false;
            }

            parsedTable = fallbackTable;
        }

        table = parsedTable;
        if (!attributeSet.IsEmpty && table.Attributes.IsEmpty) {
            table.SetAttributes(attributeSet);
            MarkdownGenericAttributeSourceSpans.Set(table, attributeSourceText, attributeSourceSpan);
        }

        index = j;
        return true;
    }

    private static bool TryParseCollectedNestedBlock<TBlock>(
        List<string> lines,
        MarkdownReaderOptions options,
        MarkdownReaderState state,
        int lineOffset,
        out TBlock? block)
        where TBlock : class, IMarkdownBlock {
        block = null;
        if (lines == null || lines.Count == 0) return false;

        var nested = ParseBlocksFromLines(lines.ToArray(), options, state, lineOffset: lineOffset);
        if (nested.Count == 0 || nested[0] is not TBlock parsedBlock) {
            return false;
        }

        block = parsedBlock;
        return true;
    }

    private static bool TryParseNestedHtmlBlock(string[] lines, ref int index, int continuationIndent, MarkdownReaderOptions options, MarkdownReaderState state, out IMarkdownBlock? block) {
        block = null;
        if (lines == null || index < 0 || index >= lines.Length) return false;
        if (!options.HtmlBlocks) return false;

        string line = lines[index] ?? string.Empty;
        if (CountLeadingIndentColumns(line) < continuationIndent) return false;
        string slice = StripLeadingIndentColumns(line, continuationIndent);
        string sliceTrim = slice.TrimStart();
        if (!sliceTrim.StartsWith("<")) return false;
        if (TryParseAngleAutolink(sliceTrim, 0, out _, out _, out _)) return false;

        // Collect contiguous indented lines and let HtmlBlockParser decide the extent.
        int j = index;
        var collected = new List<string>();
        while (j < lines.Length) {
            string raw = lines[j] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(raw)) {
                // Allow unindented blank lines inside HTML blocks within list items.
                collected.Add(string.Empty);
                j++;
                continue;
            }
            if (CountLeadingIndentColumns(raw) < continuationIndent) break;
            collected.Add(StripLeadingIndentColumns(raw, continuationIndent));
            j++;
        }
        if (collected.Count == 0) return false;

        int local = 0;
        var tempDoc = MarkdownDoc.Create();
        var parser = new HtmlBlockParser();
        if (!parser.TryParse(collected.ToArray(), ref local, options, tempDoc, state)) return false;
        if (tempDoc.Blocks.Count != 1) return false;

        block = tempDoc.Blocks[0];
        index = index + local;
        return true;
    }

}

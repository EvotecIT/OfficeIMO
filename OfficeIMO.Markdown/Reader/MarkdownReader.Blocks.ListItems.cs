using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static void ConsumeNestedBlocksForListItem(
        string[] lines,
        ref int index,
        int itemLevelAbs,
        int continuationIndent,
        MarkdownReaderOptions options,
        MarkdownReaderState state,
        ListItem item,
        bool allowNestedOrdered,
        bool allowNestedUnordered) {

        if (lines == null || item == null) return;

        while (index < lines.Length) {
            if (!state.IsListBoundaryProbe && IsStructurallyBlankListItem(item) && string.IsNullOrWhiteSpace(lines[index])) {
                return;
            }

            int k = index;
            bool sawBlankLine = false;

            // Skip blank lines only when they are followed by nested content.
            while (k < lines.Length && string.IsNullOrWhiteSpace(lines[k])) {
                sawBlankLine = true;
                int peek = k + 1;
                if (peek >= lines.Length) return;
                var next = lines[peek] ?? string.Empty;
                if (string.IsNullOrWhiteSpace(next)) {
                    k = peek;
                    continue;
                }
                if (CountLeadingIndentColumns(next) < continuationIndent) return;
                if (!IsListNestedBlockStart(next, continuationIndent, itemLevelAbs, allowNestedOrdered, allowNestedUnordered, options)) {
                    k = peek;
                    break;
                }
                k = peek;
            }

            if (k >= lines.Length) { index = k; return; }
            if (!sawBlankLine && k > 0 && string.IsNullOrWhiteSpace(lines[k - 1])) sawBlankLine = true;

            // Nested fenced code block
            int tmp = k;
            if (TryParseNestedFencedCodeBlock(lines, ref tmp, continuationIndent, options, state, out var code) && code != null) {
                item.NestedBlocks.Add(code);
                AddListItemChildSyntaxNode(item, code, lines, continuationIndent, k, tmp, state);
                if (sawBlankLine) item.ForceLoose = true;
                index = tmp;
                continue;
            }

            // Nested indented code block
            tmp = k;
            if (TryParseNestedIndentedCodeBlock(lines, ref tmp, continuationIndent, options, out var indented) && indented != null) {
                item.NestedBlocks.Add(indented);
                AddListItemChildSyntaxNode(item, indented, k, tmp, state);
                if (sawBlankLine) item.ForceLoose = true;
                index = tmp;
                continue;
            }

            // Nested blockquote
            tmp = k;
            if (TryParseNestedQuoteBlock(lines, ref tmp, itemLevelAbs, continuationIndent, options, state, out var quote) && quote != null) {
                item.NestedBlocks.Add(quote);
                AddListItemChildSyntaxNode(item, quote, lines, continuationIndent, k, tmp, state);
                if (sawBlankLine) item.ForceLoose = true;
                index = tmp;
                continue;
            }

            // Nested custom container
            tmp = k;
            if (TryParseNestedCustomContainerBlock(lines, ref tmp, continuationIndent, options, state, out var customContainer, out var customContainerSyntaxNode) && customContainer != null) {
                item.NestedBlocks.Add(customContainer);
                if (customContainerSyntaxNode != null) {
                    item.SyntaxChildren.Add(customContainerSyntaxNode);
                } else {
                    AddListItemChildSyntaxNode(item, customContainer, lines, continuationIndent, k, tmp, state);
                }
                if (sawBlankLine) item.ForceLoose = true;
                index = tmp;
                continue;
            }

            // Nested table
            tmp = k;
            if (TryParseNestedTableBlock(lines, ref tmp, continuationIndent, options, state, out var table) && table != null) {
                item.NestedBlocks.Add(table);
                AddListItemChildSyntaxNode(item, table, k, tmp, state);
                if (sawBlankLine) item.ForceLoose = true;
                index = tmp;
                continue;
            }

            // Nested HTML blocks (details / raw HTML) when HtmlBlocks are enabled.
            tmp = k;
            if (TryParseNestedHtmlBlock(lines, ref tmp, continuationIndent, options, state, out var htmlBlock) && htmlBlock != null) {
                item.NestedBlocks.Add(htmlBlock);
                AddListItemChildSyntaxNode(item, htmlBlock, k, tmp, state);
                if (sawBlankLine) item.ForceLoose = true;
                index = tmp;
                continue;
            }

            if (TryParseNestedAttributedListBlock(
                    lines,
                    k,
                    itemLevelAbs,
                    continuationIndent,
                    options,
                    state,
                    allowNestedOrdered,
                    allowNestedUnordered,
                    out var attributedList,
                    out var attributedListEndIndex)) {
                item.NestedBlocks.Add(attributedList);
                AddListItemChildSyntaxNode(item, attributedList, lines, continuationIndent, k, attributedListEndIndex, state);
                if (sawBlankLine) item.ForceLoose = true;
                index = attributedListEndIndex;
                continue;
            }

            // Nested ordered list
            if (allowNestedOrdered
                && options.OrderedLists
                && CountLeadingIndentColumns(lines[k] ?? string.Empty) >= continuationIndent
                && IsOrderedListLine(lines[k], options, out int lvlAbsO2, out _, out _, out _)
                && lvlAbsO2 >= itemLevelAbs + 1) {
                if (TryParseNestedListBlock(lines, k, continuationIndent, options, state, new OrderedListParser(), out var orderedList, out var orderedEndIndex, out var orderedSyntaxNode)) {
                    item.NestedBlocks.Add(orderedList);
                    if (orderedSyntaxNode != null) {
                        item.SyntaxChildren.Add(orderedSyntaxNode);
                    } else {
                        AddListItemChildSyntaxNode(item, orderedList, lines, continuationIndent, k, orderedEndIndex, state);
                    }
                    if (sawBlankLine) item.ForceLoose = true;
                    index = orderedEndIndex;
                    continue;
                }
            }

            // Nested unordered list
            if (allowNestedUnordered
                && options.UnorderedLists
                && CountLeadingIndentColumns(lines[k] ?? string.Empty) >= continuationIndent
                && IsUnorderedListLine(lines[k], out int lvlAbsU2)
                && lvlAbsU2 >= itemLevelAbs + 1) {
                if (TryParseNestedListBlock(lines, k, continuationIndent, options, state, new UnorderedListParser(), out var unorderedList, out var unorderedEndIndex, out var unorderedSyntaxNode)) {
                    item.NestedBlocks.Add(unorderedList);
                    if (unorderedSyntaxNode != null) {
                        item.SyntaxChildren.Add(unorderedSyntaxNode);
                    } else {
                        AddListItemChildSyntaxNode(item, unorderedList, lines, continuationIndent, k, unorderedEndIndex, state);
                    }
                    if (sawBlankLine) item.ForceLoose = true;
                    index = unorderedEndIndex;
                    continue;
                }
            }

            tmp = k;
            if (TryParseTrailingParagraphsForListItem(lines, ref tmp, itemLevelAbs, continuationIndent, options, state, out var trailingParagraphs, out var trailingSyntaxNodes) && trailingParagraphs.Count > 0) {
                foreach (var paragraph in trailingParagraphs) {
                    item.NestedBlocks.Add(paragraph);
                }
                for (int p = 0; p < trailingSyntaxNodes.Count; p++) {
                    item.SyntaxChildren.Add(trailingSyntaxNodes[p]);
                }
                if (sawBlankLine || item.NestedBlocks.Count > 0) item.ForceLoose = true;
                index = tmp;
                continue;
            }

            // Nothing nested to consume.
            index = k;
            return;
        }
    }

    private static bool IsStructurallyBlankListItem(ListItem item) {
        return item.Content.Nodes.Count == 0
               && !item.HasAdditionalParagraphs
               && item.NestedBlockCount == 0;
    }

    private static bool TryParseNestedListBlock(
        string[] lines,
        int startIndex,
        int continuationIndent,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        IMarkdownBlockParser parser,
        out IMarkdownListBlock list,
        out int endIndex,
        out MarkdownSyntaxNode? syntaxNode) {
        var effectiveState = state ?? new MarkdownReaderState();
        if (effectiveState.CaptureSyntaxTree) {
            return TryParseSourceBackedNestedListBlock(
                lines,
                startIndex,
                continuationIndent,
                options,
                effectiveState,
                parser,
                out list,
                out endIndex,
                out syntaxNode);
        }

        int idx = startIndex;
        var tempDoc = MarkdownDoc.Create();
        int previousMarkerIndentOffset = effectiveState.ListMarkerIndentOffset;
        effectiveState.ListMarkerIndentOffset = CountLeadingIndentColumns(lines[startIndex] ?? string.Empty);
        try {
            if (parser.TryParse(lines, ref idx, options, tempDoc, effectiveState) &&
                tempDoc.Blocks.Count == 1 &&
                tempDoc.Blocks[0] is IMarkdownListBlock parsed) {
                list = parsed;
                endIndex = idx;
                syntaxNode = null;
                return true;
            }
        } finally {
            effectiveState.ListMarkerIndentOffset = previousMarkerIndentOffset;
        }

        list = null!;
        endIndex = startIndex;
        syntaxNode = null;
        return false;
    }

    private static bool TryParseSourceBackedNestedListBlock(
        string[] lines,
        int startIndex,
        int continuationIndent,
        MarkdownReaderOptions options,
        MarkdownReaderState state,
        IMarkdownBlockParser parser,
        out IMarkdownListBlock list,
        out int endIndex,
        out MarkdownSyntaxNode? syntaxNode) {
        int idx = startIndex;
        var nestedDocument = MarkdownDoc.Create();
        int previousMarkerIndentOffset = state.ListMarkerIndentOffset;
        bool previousRelativeListMarkerIndent = state.UseRelativeListMarkerIndent;
        bool requiresSourceRemap = false;
        state.ListMarkerIndentOffset = CountLeadingIndentColumns(lines[startIndex] ?? string.Empty);
        state.UseRelativeListMarkerIndent = true;
        try {
            bool parsedNestedList;
            using (nestedDocument.DeferObjectTreeBinding(completeBindingOnDispose: false)) {
                parsedNestedList = parser.TryParse(lines, ref idx, options, nestedDocument, state);
            }

            if (parsedNestedList &&
                nestedDocument.Blocks.Count == 1 &&
                nestedDocument.Blocks[0] is IMarkdownListBlock parsed) {
                requiresSourceRemap = state.SourceLineAbsoluteNumbers != null ||
                    !IsDirectSourceBackedNestedList(lines, startIndex, idx, options);
                if (!requiresSourceRemap) {
                    var parsedObject = (MarkdownObject)parsed;
                    var firstLine = lines[startIndex] ?? string.Empty;
                    var lastLine = lines[Math.Max(startIndex, idx - 1)] ?? string.Empty;
                    parsedObject.SourceSpan = CreateSpan(
                        state,
                        state.SourceLineOffset + startIndex + 1,
                        CountLeadingIndentColumns(firstLine) + 1,
                        state.SourceLineOffset + Math.Max(startIndex + 1, idx),
                        Math.Max(1, lastLine.Length));

                    list = parsed;
                    endIndex = idx;
                    // Preserve the child-cache position with a lightweight source root. The
                    // enclosing builder expands it once from the parsed list items instead of
                    // building and recursively cloning an identical intermediate subtree here.
                    syntaxNode = new MarkdownSyntaxNode(
                        parsed.ListSyntaxKind,
                        parsedObject.SourceSpan,
                        parsed.ListLiteral,
                        associatedObject: parsed);
                    return true;
                }
            }
        } finally {
            state.ListMarkerIndentOffset = previousMarkerIndentOffset;
            state.UseRelativeListMarkerIndent = previousRelativeListMarkerIndent;
        }

        if (requiresSourceRemap) {
            var slices = BuildListItemNestedSourceLines(lines, continuationIndent, startIndex, idx, state);
            var (blocks, syntaxChildren) = ParseNestedMarkdownBlocks(slices, options, state);
            if (blocks.Count == 1 &&
                blocks[0] is IMarkdownListBlock sourceMappedList &&
                syntaxChildren.Count == 1) {
                list = sourceMappedList;
                endIndex = idx;
                syntaxNode = syntaxChildren[0];
                return true;
            }
        }

        list = null!;
        endIndex = startIndex;
        syntaxNode = null;
        return false;
    }

    private static bool IsDirectSourceBackedNestedList(
        string[] lines,
        int startIndex,
        int endIndex,
        MarkdownReaderOptions options) {
        for (int lineIndex = startIndex; lineIndex < endIndex; lineIndex++) {
            var line = lines[lineIndex] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(line) || line.IndexOf('\t') >= 0) {
                return false;
            }

            if (!IsUnorderedListLine(line) &&
                !IsOrderedListLine(line, options, out _, out _, out _, out _)) {
                return false;
            }
        }

        return endIndex > startIndex;
    }

    private static bool TryParseNestedAttributedListBlock(
        string[] lines,
        int attributeLineIndex,
        int itemLevelAbs,
        int continuationIndent,
        MarkdownReaderOptions options,
        MarkdownReaderState state,
        bool allowNestedOrdered,
        bool allowNestedUnordered,
        out IMarkdownListBlock list,
        out int endIndex) {

        list = null!;
        endIndex = attributeLineIndex;

        if (!TryConsumeNestedStandaloneGenericAttributeLine(
                lines,
                attributeLineIndex,
                continuationIndent,
                options,
                state,
                NestedStandaloneGenericAttributeTarget.List,
                out var attributeSet,
                out var attributeSourceText,
                out var attributeSourceSpan)) {
            return false;
        }

        int listStartIndex = attributeLineIndex + 1;
        if (listStartIndex >= lines.Length) {
            return false;
        }

        IMarkdownBlockParser? parser = null;
        var listLine = lines[listStartIndex] ?? string.Empty;
        if (allowNestedOrdered
            && options.OrderedLists
            && CountLeadingIndentColumns(listLine) >= continuationIndent
            && IsOrderedListLine(listLine, options, out int orderedLevel, out _, out _, out _)
            && orderedLevel >= itemLevelAbs + 1) {
            parser = new OrderedListParser();
        } else if (allowNestedUnordered
            && options.UnorderedLists
            && CountLeadingIndentColumns(listLine) >= continuationIndent
            && IsUnorderedListLine(listLine, out int unorderedLevel)
            && unorderedLevel >= itemLevelAbs + 1) {
            parser = new UnorderedListParser();
        }

        if (parser == null
            || !TryParseNestedListBlock(lines, listStartIndex, continuationIndent, options, state, parser, out var parsedList, out var parsedEndIndex, out _)) {
            return false;
        }

        if (parsedList is MarkdownObject markdownObject && markdownObject.Attributes.IsEmpty) {
            markdownObject.SetAttributes(attributeSet);
            MarkdownGenericAttributeSourceSpans.Set(markdownObject, attributeSourceText, attributeSourceSpan);
        }

        list = parsedList;
        endIndex = parsedEndIndex;
        return true;
    }

    private static bool TryParseTrailingParagraphsForListItem(
        string[] lines,
        ref int index,
        int itemLevelAbs,
        int continuationIndent,
        MarkdownReaderOptions options,
        MarkdownReaderState state,
        out List<ParagraphBlock> paragraphs,
        out List<MarkdownSyntaxNode> syntaxNodes) {

        paragraphs = null!;
        syntaxNodes = null!;
        if (lines == null || index < 0 || index >= lines.Length) return false;

        string line = lines[index] ?? string.Empty;
        if (string.IsNullOrWhiteSpace(line)) return false;
        if (CountLeadingIndentColumns(line) < continuationIndent) return false;
        if (IsListNestedBlockStart(line, continuationIndent, itemLevelAbs, allowNestedOrdered: true, allowNestedUnordered: true, options)) return false;
        if (IsUnorderedListLine(line) || IsOrderedListLine(line, options, out _, out _, out _, out _)) return false;

        var attributeSourceText = string.Empty;
        MarkdownAttributeSet attributeSet = MarkdownAttributeSet.Empty;
        MarkdownSourceSpan? attributeSourceSpan = null;
        if (TryConsumeNestedStandaloneGenericAttributeLine(
                lines,
                index,
                continuationIndent,
                options,
                state,
                NestedStandaloneGenericAttributeTarget.Paragraph,
                out attributeSet,
                out attributeSourceText,
                out attributeSourceSpan)) {
            index++;
            if (index >= lines.Length) {
                return false;
            }

            line = lines[index] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(line)) return false;
            if (CountLeadingIndentColumns(line) < continuationIndent) return false;
            if (IsListNestedBlockStart(line, continuationIndent, itemLevelAbs, allowNestedOrdered: true, allowNestedUnordered: true, options)) return false;
            if (IsUnorderedListLine(line) || IsOrderedListLine(line, options, out _, out _, out _, out _)) return false;
        }

        paragraphs = new List<ParagraphBlock>();
        syntaxNodes = new List<MarkdownSyntaxNode>();

        string firstContent = StripLeadingIndentColumns(line, continuationIndent);
        firstContent = firstContent.TrimStart();
        int firstStartColumn = continuationIndent + CountLeadingIndentColumns(StripLeadingIndentColumns(line, continuationIndent)) + 1;

        int next = index + 1;
        var paragraphSourceLines = new List<MarkdownSourceLineSlice>();
        var paragraphLines = ConsumeListContinuationLines(
            lines,
            ref next,
            continuationIndent,
            firstContent,
            options,
            sourceLines: paragraphSourceLines,
            absoluteLineOffset: state.SourceLineOffset,
            initialLineIndex: index,
            initialStartColumn: firstStartColumn,
            state: state);
        paragraphs.AddRange(ParseParagraphBlocksFromSourceLines(paragraphSourceLines, options, state));
        if (!attributeSet.IsEmpty && paragraphs.Count > 0) {
            paragraphs[0].SetAttributes(attributeSet);
            MarkdownGenericAttributeSourceSpans.Set(paragraphs[0], attributeSourceText, attributeSourceSpan);
        }

        AddParagraphSyntaxNodes(syntaxNodes, paragraphSourceLines, options, state);

        index = next;
        return paragraphs.Count > 0;
    }

    private static bool IsListNestedBlockStart(
        string line,
        int continuationIndent,
        int itemLevelAbs,
        bool allowNestedOrdered,
        bool allowNestedUnordered,
        MarkdownReaderOptions options) {

        if (string.IsNullOrEmpty(line)) return false;

        int nextIndentColumns = CountLeadingIndentColumns(line);
        if (nextIndentColumns < continuationIndent) return false;

        if (allowNestedOrdered && options.OrderedLists &&
            IsOrderedListLine(line, options, out int lvlAbsO, out _, out _, out _) &&
            lvlAbsO >= itemLevelAbs + 1) {
            return true;
        }

        if (allowNestedUnordered && options.UnorderedLists &&
            IsUnorderedListLine(line, out int lvlAbsU) &&
            lvlAbsU >= itemLevelAbs + 1) {
            return true;
        }

        var slice = StripLeadingIndentColumns(line, continuationIndent);
        var sliceTrim = slice.TrimStart();

        if (options.FencedCode && IsCodeFenceOpen(slice, out _, out _, out _)) return true;
        if (options.IndentedCodeBlocks && nextIndentColumns >= continuationIndent + 4 && !string.IsNullOrWhiteSpace(slice)) return true;
        if (sliceTrim.StartsWith(">", StringComparison.Ordinal)) return true;

        if (options.Tables && LooksLikeTableRow(sliceTrim)) return true;

        if (options.HtmlBlocks && sliceTrim.StartsWith("<", StringComparison.Ordinal) && !TryParseAngleAutolink(sliceTrim, 0, out _, out _, out _)) {
            return true;
        }

        return false;
    }

    private static bool IsOrderedListLine(string line, out int number, out string content) {
        return IsOrderedListLine(line, null, out number, out content);
    }

    private static bool IsOrderedListLine(string line, MarkdownReaderOptions? options, out int number, out string content) {
        number = 0;
        content = string.Empty;
        if (!TryGetOrderedListMarkerInfo(line, options, out _, out number, out int contentStartIndex, out _, out _)) return false;
        content = line.Substring(contentStartIndex);
        return true;
    }

    private static bool IsOrderedListLine(string line, out int level, out int number, out string content) {
        return IsOrderedListLine(line, null, out level, out number, out content, out _);
    }

    private static bool IsOrderedListLine(
        string line,
        MarkdownReaderOptions? options,
        out int level,
        out int number,
        out string content,
        out MarkdownOrderedListMarkerStyle markerStyle) {
        level = 0;
        number = 0;
        content = string.Empty;
        markerStyle = MarkdownOrderedListMarkerStyle.Decimal;
        if (!TryGetOrderedListMarkerInfo(line, options, out int spaces, out number, out int contentStartIndex, out _, out markerStyle)) return false;
        content = line.Substring(contentStartIndex);
        level = spaces / 2;
        return true;
    }

    private static bool IsUnorderedListLine(string line, out bool isTask, out bool done, out string content) {
        isTask = false;
        done = false;
        content = string.Empty;
        if (!TryGetUnorderedListMarkerInfo(line, out _, out int contentStartIndex)) return false;

        var c = line.Substring(contentStartIndex);
        if (TryStripTaskListMarker(c, out done, out content)) {
            isTask = true;
            return true;
        }

        content = c;
        return true;
    }

    private static bool IsUnorderedListLine(string line) =>
        TryGetUnorderedListMarkerInfo(line, out _, out _);

    private static bool IsUnorderedListLine(string line, out int level) {
        level = 0;
        if (!TryGetUnorderedListMarkerInfo(line, out int spaces, out _)) {
            return false;
        }

        level = spaces / 2;
        return true;
    }

    private static bool IsUnorderedListLine(string line, bool parseTaskMarker, out int level, out bool isTask, out bool done, out string content) {
        level = 0;
        isTask = false;
        done = false;
        content = string.Empty;
        if (!TryGetUnorderedListMarkerInfo(line, out int spaces, out int contentStartIndex)) return false;

        if (parseTaskMarker
            && TryGetTaskListMarkerContentStartIndex(line, contentStartIndex, out int taskContentStartIndex)) {
            isTask = true;
            done = line[contentStartIndex + 1] is 'x' or 'X';
            content = taskContentStartIndex >= line.Length
                ? string.Empty
                : line.Substring(taskContentStartIndex);
            level = spaces / 2;
            return true;
        }

        content = line.Substring(contentStartIndex);
        level = spaces / 2;
        return true;
    }

    private static string GetUnorderedListItemContent(string line) {
        return TryGetUnorderedListMarkerInfo(line, out _, out int contentStartIndex)
            ? line.Substring(contentStartIndex)
            : string.Empty;
    }

    private static bool IsCalloutHeader(string line, out string kind, out string title) {
        kind = string.Empty; title = string.Empty;
        if (string.IsNullOrEmpty(line)) return false;
        var t = line.TrimStart();
        if (!t.StartsWith(">", StringComparison.Ordinal)) return false;
        t = t.Substring(1).TrimStart();
        if (!t.StartsWith("[!", StringComparison.Ordinal)) return false;
        int close = t.IndexOf(']');
        if (close < 0 || close < 3) return false;
        string marker = t.Substring(2, close - 2);
        for (int i = 0; i < marker.Length; i++) if (!char.IsLetter(marker[i])) return false;
        kind = marker.ToLowerInvariant();
        title = t.Substring(close + 1).TrimStart();
        // Title is optional: "> [!NOTE]" is valid and should produce a callout with the default title for the kind.
        return true;
    }

    private static bool IsCalloutHeader(string line, MarkdownReaderOptions options, out string kind, out string title) {
        if (!IsCalloutHeader(line, out kind, out title)) return false;
        return options.CalloutTitleMode != MarkdownCalloutTitleMode.MarkdigCompatible || string.IsNullOrWhiteSpace(title);
    }

}

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
            if (IsStructurallyBlankListItem(item) && string.IsNullOrWhiteSpace(lines[index])) {
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
                item.Children.Add(code);
                AddListItemChildSyntaxNode(item, code, lines, continuationIndent, k, tmp, state);
                if (sawBlankLine) item.ForceLoose = true;
                index = tmp;
                continue;
            }

            // Nested indented code block
            tmp = k;
            if (TryParseNestedIndentedCodeBlock(lines, ref tmp, continuationIndent, options, out var indented) && indented != null) {
                item.Children.Add(indented);
                AddListItemChildSyntaxNode(item, indented, k, tmp, state);
                if (sawBlankLine) item.ForceLoose = true;
                index = tmp;
                continue;
            }

            // Nested blockquote
            tmp = k;
            if (TryParseNestedQuoteBlock(lines, ref tmp, itemLevelAbs, continuationIndent, options, state, out var quote) && quote != null) {
                item.Children.Add(quote);
                AddListItemChildSyntaxNode(item, quote, lines, continuationIndent, k, tmp, state);
                if (sawBlankLine) item.ForceLoose = true;
                index = tmp;
                continue;
            }

            // Nested custom container
            tmp = k;
            if (TryParseNestedCustomContainerBlock(lines, ref tmp, continuationIndent, options, state, out var customContainer, out var customContainerSyntaxNode) && customContainer != null) {
                item.Children.Add(customContainer);
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
                item.Children.Add(table);
                AddListItemChildSyntaxNode(item, table, k, tmp, state);
                if (sawBlankLine) item.ForceLoose = true;
                index = tmp;
                continue;
            }

            // Nested HTML blocks (details / raw HTML) when HtmlBlocks are enabled.
            tmp = k;
            if (TryParseNestedHtmlBlock(lines, ref tmp, continuationIndent, options, state, out var htmlBlock) && htmlBlock != null) {
                item.Children.Add(htmlBlock);
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
                item.Children.Add(attributedList);
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
                    item.Children.Add(orderedList);
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
                && IsUnorderedListLine(lines[k], out int lvlAbsU2, out _, out _, out _)
                && lvlAbsU2 >= itemLevelAbs + 1) {
                if (TryParseNestedListBlock(lines, k, continuationIndent, options, state, new UnorderedListParser(), out var unorderedList, out var unorderedEndIndex, out var unorderedSyntaxNode)) {
                    item.Children.Add(unorderedList);
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
                    item.Children.Add(paragraph);
                }
                for (int p = 0; p < trailingSyntaxNodes.Count; p++) {
                    item.SyntaxChildren.Add(trailingSyntaxNodes[p]);
                }
                if (sawBlankLine || item.Children.Count > 0) item.ForceLoose = true;
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
               && item.AdditionalParagraphs.Count == 0
               && item.Children.Count == 0;
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
        int idx = startIndex;
        var tempDoc = MarkdownDoc.Create();
        var effectiveState = state ?? new MarkdownReaderState();
        int previousMarkerIndentOffset = effectiveState.ListMarkerIndentOffset;
        effectiveState.ListMarkerIndentOffset = CountLeadingIndentColumns(lines[startIndex] ?? string.Empty);
        try {
            if (parser.TryParse(lines, ref idx, options, tempDoc, effectiveState) &&
                tempDoc.Blocks.Count == 1 &&
                tempDoc.Blocks[0] is IMarkdownListBlock parsedList) {
                var slices = BuildListItemNestedSourceLines(lines, continuationIndent, startIndex, idx, effectiveState);
                var (blocks, syntaxChildren) = ParseNestedMarkdownBlocks(slices, options, effectiveState);
                if (blocks.Count == 1 &&
                    blocks[0] is IMarkdownListBlock sourceMappedList &&
                    syntaxChildren.Count == 1) {
                    list = sourceMappedList;
                    syntaxNode = syntaxChildren[0];
                    endIndex = idx;
                    return true;
                }

                list = parsedList;
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
            && IsUnorderedListLine(listLine, out int unorderedLevel, out _, out _, out _)
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

        paragraphs = new List<ParagraphBlock>();
        syntaxNodes = new List<MarkdownSyntaxNode>();
        if (lines == null || index < 0 || index >= lines.Length) return false;

        string line = lines[index] ?? string.Empty;
        if (string.IsNullOrWhiteSpace(line)) return false;
        if (CountLeadingIndentColumns(line) < continuationIndent) return false;
        if (IsListNestedBlockStart(line, continuationIndent, itemLevelAbs, allowNestedOrdered: true, allowNestedUnordered: true, options)) return false;
        if (IsUnorderedListLine(line, out _, out _, out _, out _) || IsOrderedListLine(line, options, out _, out _, out _, out _)) return false;

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
            IsUnorderedListLine(line, out int lvlAbsU, out _, out _, out _) &&
            lvlAbsU >= itemLevelAbs + 1) {
            return true;
        }

        var slice = StripLeadingIndentColumns(line, continuationIndent);
        var sliceTrim = slice.TrimStart();

        if (options.FencedCode && IsCodeFenceOpen(slice, out _, out _, out _)) return true;
        if (options.IndentedCodeBlocks && nextIndentColumns >= continuationIndent + 4 && !string.IsNullOrWhiteSpace(slice)) return true;
        if (sliceTrim.StartsWith(">")) return true;

        if (options.Tables && LooksLikeTableRow(sliceTrim)) return true;

        if (options.HtmlBlocks && sliceTrim.StartsWith("<") && !TryParseAngleAutolink(sliceTrim, 0, out _, out _, out _)) {
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

    private static bool IsUnorderedListLine(string line, out int level, out bool isTask, out bool done, out string content) {
        level = 0;
        isTask = false;
        done = false;
        content = string.Empty;
        if (!TryGetUnorderedListMarkerInfo(line, out int spaces, out int contentStartIndex)) return false;

        string c = line.Substring(contentStartIndex);
        if (TryStripTaskListMarker(c, out done, out content)) {
            isTask = true;
            level = spaces / 2;
            return true;
        }

        content = c;
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
        if (!t.StartsWith(">")) return false;
        t = t.Substring(1).TrimStart();
        if (!t.StartsWith("[!")) return false;
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

    private static int GetListContinuationIndent(string line, MarkdownReaderOptions? options = null) {
        if (string.IsNullOrEmpty(line)) return 0;
        if (TryGetOrderedListMarkerInfo(line, options, out int orderedLeadingSpaces, out _, out int orderedContentStartIndex, out _, out _)) {
            if (string.IsNullOrWhiteSpace(line.Substring(orderedContentStartIndex))
                && TryGetOrderedListMarkerWidth(line, orderedLeadingSpaces, options, out int orderedMarkerWidth)) {
                return orderedLeadingSpaces + orderedMarkerWidth + 1;
            }

            return orderedContentStartIndex;
        }

        if (TryGetUnorderedListMarkerInfo(line, out int unorderedLeadingSpaces, out int unorderedContentStartIndex)) {
            if (string.IsNullOrWhiteSpace(line.Substring(unorderedContentStartIndex))) {
                return unorderedLeadingSpaces + 2;
            }

            return unorderedContentStartIndex;
        }

        int spaces = CountLeadingSpaces(line);
        return spaces + 2;
    }

    private static int GetRelativeListItemLevel(List<int>? continuationIndentsByLevel, string line) {
        if (continuationIndentsByLevel == null || continuationIndentsByLevel.Count == 0 || string.IsNullOrEmpty(line)) {
            return 0;
        }

        int indentColumns = CountLeadingIndentColumns(line);
        for (int level = continuationIndentsByLevel.Count - 1; level >= 0; level--) {
            if (indentColumns >= continuationIndentsByLevel[level]) {
                return level + 1;
            }
        }

        return 0;
    }

    private static void TrackListItemContinuationIndent(List<int> continuationIndentsByLevel, int level, int continuationIndent) {
        if (continuationIndentsByLevel == null) {
            return;
        }

        while (continuationIndentsByLevel.Count > level) {
            continuationIndentsByLevel.RemoveAt(continuationIndentsByLevel.Count - 1);
        }

        if (continuationIndentsByLevel.Count == level) {
            continuationIndentsByLevel.Add(continuationIndent);
            return;
        }

        continuationIndentsByLevel[level] = continuationIndent;
    }

    private static int GetTaskMarkerConsumedColumns(string content) {
        if (string.IsNullOrEmpty(content)) return 0;
        return TryGetTaskListMarkerContentStartIndex(content, out int contentStartIndex)
            ? contentStartIndex
            : 0;
    }

    private static void SetListItemMarkerSourceSpans(ListItem item, string line, int lineIndex, bool isTask, MarkdownReaderOptions options, MarkdownReaderState state) {
        if (item == null) {
            return;
        }

        int absoluteLineNumber = state.SourceLineOffset + lineIndex + 1;
        item.MarkerSourceSpan = TryCreateListMarkerSourceSpan(line, absoluteLineNumber, options, state);
        item.MarkerText = TryGetListMarkerText(line, options);
        if (isTask) {
            item.TaskMarkerSourceSpan = TryCreateTaskMarkerSourceSpan(line, absoluteLineNumber, options, state);
            item.TaskMarkerText = TryGetTaskMarkerText(line);
        } else {
            item.TaskMarkerText = null;
        }
    }

    private static string? TryGetListMarkerText(string line, MarkdownReaderOptions? options = null) {
        if (string.IsNullOrEmpty(line)) {
            return null;
        }

        if (TryGetOrderedListMarkerInfo(line, options, out int orderedLeadingSpaces, out _, out _, out _, out _)) {
            int delimiterIndex = GetOrderedListMarkerDelimiterIndex(line, orderedLeadingSpaces);
            return delimiterIndex < line.Length
                ? line.Substring(orderedLeadingSpaces, delimiterIndex - orderedLeadingSpaces + 1)
                : null;
        }

        return TryGetUnorderedListMarkerInfo(line, out int unorderedLeadingSpaces, out _, out char marker)
            ? marker.ToString()
            : null;
    }

    private static string? TryGetTaskMarkerText(string line) {
        if (string.IsNullOrEmpty(line)
            || !TryGetRawListItemContentAfterMarker(line, out string content)
            || !TryGetTaskListMarkerContentStartIndex(content, out _)) {
            return null;
        }

        return content.Length >= 3 ? content.Substring(0, 3) : null;
    }

    private static MarkdownSourceSpan? TryCreateListMarkerSourceSpan(string line, int absoluteLineNumber, MarkdownReaderOptions? options, MarkdownReaderState state) {
        if (string.IsNullOrEmpty(line)) {
            return null;
        }

        if (TryGetOrderedListMarkerInfo(line, options, out int orderedLeadingSpaces, out _, out _, out _, out _)) {
            int delimiterIndex = GetOrderedListMarkerDelimiterIndex(line, orderedLeadingSpaces);
            if (delimiterIndex < line.Length) {
                return CreateSpan(state, absoluteLineNumber, orderedLeadingSpaces + 1, absoluteLineNumber, delimiterIndex + 1);
            }
        }

        if (TryGetUnorderedListMarkerInfo(line, out int unorderedLeadingSpaces, out _, out _)) {
            int markerColumn = unorderedLeadingSpaces + 1;
            return CreateSpan(state, absoluteLineNumber, markerColumn, absoluteLineNumber, markerColumn);
        }

        return null;
    }

    private static MarkdownSourceSpan? TryCreateTaskMarkerSourceSpan(string line, int absoluteLineNumber, MarkdownReaderOptions options, MarkdownReaderState state) {
        if (string.IsNullOrEmpty(line)) {
            return null;
        }

        if (!TryGetRawListItemContentAfterMarker(line, out string content)
            || !TryGetTaskListMarkerContentStartIndex(content, out _)) {
            return null;
        }

        int startColumn = GetListContinuationIndent(line, options) + 1;
        return CreateSpan(state, absoluteLineNumber, startColumn, absoluteLineNumber, startColumn + 2);
    }

    private static bool TryStripTaskListMarker(string content, out bool done, out string stripped) {
        done = false;
        stripped = content ?? string.Empty;
        if (!TryGetTaskListMarkerContentStartIndex(stripped, out int contentStartIndex)) {
            return false;
        }

        done = stripped[1] == 'x' || stripped[1] == 'X';
        stripped = contentStartIndex >= stripped.Length ? string.Empty : stripped.Substring(contentStartIndex);
        return true;
    }

    private static bool TryGetTaskListMarkerContentStartIndex(string content, out int contentStartIndex) {
        contentStartIndex = 0;
        if (string.IsNullOrEmpty(content) || content.Length < 3) {
            return false;
        }

        if (content[0] != '[' || content[2] != ']') {
            return false;
        }

        if (content[1] != ' ' && content[1] != 'x' && content[1] != 'X') {
            return false;
        }

        if (content.Length == 3) {
            contentStartIndex = 3;
            return true;
        }

        if (!char.IsWhiteSpace(content[3])) {
            return false;
        }

        contentStartIndex = 4;
        while (contentStartIndex < content.Length && char.IsWhiteSpace(content[contentStartIndex])) {
            contentStartIndex++;
        }

        return true;
    }

    private static bool TryGetRawListItemContentAfterMarker(string line, out string content) {
        content = string.Empty;
        if (string.IsNullOrEmpty(line)) return false;
        if (TryGetOrderedListMarkerInfo(line, out _, out _, out int orderedContentStartIndex)) {
            content = line.Substring(orderedContentStartIndex);
            return true;
        }

        if (TryGetUnorderedListMarkerInfo(line, out _, out int unorderedContentStartIndex)) {
            content = line.Substring(unorderedContentStartIndex);
            return true;
        }

        return false;
    }

    private static bool TryGetOrderedListMarkerInfo(string line, out int leadingSpaces, out int number, out int contentStartIndex) {
        return TryGetOrderedListMarkerInfo(line, null, out leadingSpaces, out number, out contentStartIndex, out _, out _);
    }

    private static bool TryGetOrderedListMarkerInfo(string line, out int leadingSpaces, out int number, out int contentStartIndex, out char delimiter) {
        return TryGetOrderedListMarkerInfo(line, null, out leadingSpaces, out number, out contentStartIndex, out delimiter, out _);
    }

    private static bool TryGetOrderedListMarkerInfo(
        string line,
        MarkdownReaderOptions? options,
        out int leadingSpaces,
        out int number,
        out int contentStartIndex,
        out char delimiter,
        out MarkdownOrderedListMarkerStyle markerStyle) {
        leadingSpaces = 0;
        number = 0;
        contentStartIndex = 0;
        delimiter = '\0';
        markerStyle = MarkdownOrderedListMarkerStyle.Decimal;
        if (string.IsNullOrEmpty(line)) return false;

        while (leadingSpaces < line.Length && line[leadingSpaces] == ' ') leadingSpaces++;

        int markerStart = leadingSpaces;
        int markerEnd = markerStart;
        while (markerEnd < line.Length && char.IsDigit(line[markerEnd])) markerEnd++;
        if (markerEnd > markerStart) {
            if (markerEnd - markerStart > 9) return false;
            if (markerEnd >= line.Length || (line[markerEnd] != '.' && line[markerEnd] != ')')) return false;
            delimiter = line[markerEnd];
            if (!TryGetListContentStartIndex(line, markerEnd, out contentStartIndex)) return false;
            if (!int.TryParse(line.Substring(markerStart, markerEnd - markerStart), NumberStyles.Integer, CultureInfo.InvariantCulture, out number)) number = 1;
            markerStyle = MarkdownOrderedListMarkerStyle.Decimal;
            return true;
        }

        if (options?.ListExtras != true) return false;

        while (markerEnd < line.Length && IsAsciiLetter(line[markerEnd])) markerEnd++;
        if (markerEnd == markerStart) return false;
        if (markerEnd >= line.Length || (line[markerEnd] != '.' && line[markerEnd] != ')')) return false;
        delimiter = line[markerEnd];
        if (!TryGetListContentStartIndex(line, markerEnd, out contentStartIndex)) return false;

        var marker = line.Substring(markerStart, markerEnd - markerStart);
        if (TryParseRomanListExtraMarker(marker, out number, out markerStyle)) {
            return true;
        }

        if (marker.Length == 1) {
            var ch = marker[0];
            markerStyle = char.IsUpper(ch) ? MarkdownOrderedListMarkerStyle.UpperAlpha : MarkdownOrderedListMarkerStyle.LowerAlpha;
            number = char.ToUpperInvariant(ch) - 'A' + 1;
            return number >= 1 && number <= 26;
        }

        number = 0;
        markerStyle = MarkdownOrderedListMarkerStyle.Decimal;
        return false;
    }

    private static int GetOrderedListMarkerDelimiterIndex(string line, int leadingSpaces) {
        var delimiterIndex = leadingSpaces;
        while (delimiterIndex < line.Length
               && (char.IsDigit(line[delimiterIndex]) || IsAsciiLetter(line[delimiterIndex]))) {
            delimiterIndex++;
        }

        return delimiterIndex;
    }

    private static bool TryParseRomanListExtraMarker(string marker, out int number, out MarkdownOrderedListMarkerStyle markerStyle) {
        number = 0;
        markerStyle = MarkdownOrderedListMarkerStyle.Decimal;
        if (string.IsNullOrEmpty(marker)) {
            return false;
        }

        var upper = char.IsUpper(marker[0]);
        var lower = marker.ToLowerInvariant();
        for (int i = 0; i < lower.Length; i++) {
            var ch = lower[i];
            if (ch != 'i' && ch != 'v' && ch != 'x') {
                return false;
            }
        }

        for (int value = 1; value <= 39; value++) {
            if (string.Equals(lower, FormatListExtraRoman(value), StringComparison.Ordinal)) {
                number = value;
                markerStyle = upper
                    ? MarkdownOrderedListMarkerStyle.UpperRoman
                    : MarkdownOrderedListMarkerStyle.LowerRoman;
                return true;
            }
        }

        return false;
    }

    private static string FormatListExtraRoman(int value) {
        var tens = value / 10;
        var ones = value % 10;
        return new string('x', tens) + (ones switch {
            0 => string.Empty,
            1 => "i",
            2 => "ii",
            3 => "iii",
            4 => "iv",
            5 => "v",
            6 => "vi",
            7 => "vii",
            8 => "viii",
            9 => "ix",
            _ => string.Empty
        });
    }

    private static bool TryGetUnorderedListMarkerInfo(string line, out int leadingSpaces, out int contentStartIndex) {
        return TryGetUnorderedListMarkerInfo(line, out leadingSpaces, out contentStartIndex, out _);
    }

    private static bool TryGetUnorderedListMarkerInfo(string line, out int leadingSpaces, out int contentStartIndex, out char marker) {
        leadingSpaces = 0;
        contentStartIndex = 0;
        marker = '\0';
        if (string.IsNullOrEmpty(line)) return false;

        while (leadingSpaces < line.Length && line[leadingSpaces] == ' ') leadingSpaces++;
        if (leadingSpaces >= line.Length) return false;

        marker = line[leadingSpaces];
        if (marker != '-' && marker != '*' && marker != '+') return false;
        return TryGetListContentStartIndex(line, leadingSpaces, out contentStartIndex);
    }

    private static bool TryGetListContentStartIndex(string line, int markerIndex, out int contentStartIndex) {
        contentStartIndex = 0;
        int paddingStart = markerIndex + 1;
        if (paddingStart >= line.Length) {
            contentStartIndex = line.Length;
            return true;
        }

        int paddingColumns = 0;
        int cursor = paddingStart;
        while (cursor < line.Length) {
            char ch = line[cursor];
            if (ch == ' ' && paddingColumns < 4) {
                paddingColumns++;
                cursor++;
                continue;
            }

            if (ch == '\t' && paddingColumns == 0) {
                contentStartIndex = cursor + 1;
                return true;
            }

            break;
        }

        if (cursor >= line.Length) {
            contentStartIndex = line.Length;
            return true;
        }

        if (paddingColumns == 0) return false;
        contentStartIndex = cursor;
        return true;
    }

    private static bool TryGetIndentedCodeListLead(string line, out int continuationIndent, out string content, out int startColumn) {
        continuationIndent = 0;
        content = string.Empty;
        startColumn = 1;
        if (string.IsNullOrEmpty(line)) return false;

        int leadingSpaces = 0;
        while (leadingSpaces < line.Length && line[leadingSpaces] == ' ') leadingSpaces++;
        if (leadingSpaces >= line.Length) return false;

        int markerWidth;
        if (TryGetOrderedListMarkerWidth(line, leadingSpaces, out markerWidth)) {
            if (!HasIndentedCodePaddingAfterMarker(line, leadingSpaces + markerWidth - 1)) return false;
        } else {
            char marker = line[leadingSpaces];
            if (marker != '-' && marker != '*' && marker != '+') return false;
            markerWidth = 1;
            if (!HasIndentedCodePaddingAfterMarker(line, leadingSpaces)) return false;
        }

        continuationIndent = leadingSpaces + markerWidth + 1;
        if (continuationIndent >= line.Length) return false;

        content = line.Substring(continuationIndent);
        startColumn = continuationIndent + 1;
        return CountLeadingIndentColumns(content) >= 4;
    }

    private static bool TryGetOrderedListMarkerWidth(string line, int leadingSpaces, out int markerWidth) {
        return TryGetOrderedListMarkerWidth(line, leadingSpaces, null, out markerWidth);
    }

    private static bool TryGetOrderedListMarkerWidth(string line, int leadingSpaces, MarkdownReaderOptions? options, out int markerWidth) {
        markerWidth = 0;
        if (string.IsNullOrEmpty(line) || leadingSpaces >= line.Length) return false;

        if (!TryGetOrderedListMarkerInfo(line, options, out _, out _, out _, out _, out _)) {
            return false;
        }

        var delimiterIndex = GetOrderedListMarkerDelimiterIndex(line, leadingSpaces);
        markerWidth = delimiterIndex - leadingSpaces + 1;
        return true;
    }

    private static bool HasIndentedCodePaddingAfterMarker(string line, int markerEndIndex) {
        int paddingStart = markerEndIndex + 1;
        if (paddingStart >= line.Length || line[paddingStart] != ' ') return false;

        int spaces = 0;
        int cursor = paddingStart;
        while (cursor < line.Length && line[cursor] == ' ') {
            spaces++;
            cursor++;
        }

        return spaces >= 5;
    }

    private static int GetListLeadContentStartColumn(string line, MarkdownReaderOptions? options = null, bool stripTaskMarker = false) {
        int startColumn = GetListContinuationIndent(line, options) + 1;
        if (!stripTaskMarker) return startColumn;

        return TryGetRawListItemContentAfterMarker(line, out string content)
            ? startColumn + GetTaskMarkerConsumedColumns(content)
            : startColumn;
    }
}

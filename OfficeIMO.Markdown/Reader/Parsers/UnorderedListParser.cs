namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class UnorderedListParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.UnorderedLists) return false;
            if (!IsUnorderedListLine(lines[i], out int level0Abs, out var isTask, out var done, out var firstContent)) return false;
            if (!TryGetUnorderedListMarkerInfo(lines[i], out int firstMarkerIndent, out _, out char firstMarker)) return false;
            if (options.StrictListIndentation && firstMarkerIndent - state.ListMarkerIndentOffset > 3) return false;
            using var headingAttributeScope = SuppressHeadingGenericAttributesInListItems(state);
            if (isTask && !options.TaskLists) {
                isTask = false;
                done = false;
                firstContent = GetUnorderedListItemContent(lines[i]);
            }
            var ul = new UnorderedListBlock();
            var continuationIndentsByLevel = options.StrictListIndentation ? new List<int>() : null;
            int firstContinuationIndent = GetListContinuationIndent(lines[i], options);
            int firstStartColumn;
            if (!isTask && TryGetIndentedCodeListLead(lines[i], options, out int codeLeadIndent, out string codeLeadContent, out int codeLeadStartColumn)) {
                firstContinuationIndent = codeLeadIndent;
                firstContent = codeLeadContent;
                firstStartColumn = codeLeadStartColumn;
            } else {
                firstStartColumn = GetListLeadContentStartColumn(lines[i], options, isTask);
            }

            int j = i + 1;
            var firstSourceLines = new List<MarkdownSourceLineSlice>();
            var firstLines = ConsumeListContinuationLines(
                lines,
                ref j,
                firstContinuationIndent,
                firstContent,
                options,
                breakOnAnyOrderedListLine: options.ListExtras,
                sourceLines: firstSourceLines,
                absoluteLineOffset: state.SourceLineOffset,
                initialLineIndex: i,
                initialStartColumn: firstStartColumn,
                state: state);
            var first = CreateListItemFromLeadLines(firstLines, isTask, done, options, state, i, firstSourceLines);
            first.Level = 0;
            SetListItemMarkerSourceSpans(first, lines[i], i, isTask, options, state);
            if (continuationIndentsByLevel != null) {
                TrackListItemContinuationIndent(continuationIndentsByLevel, first.Level, firstContinuationIndent);
            }
            AddListItemLeadSyntaxNodes(first, firstLines, i, options, state, firstSourceLines);
            ul.Items.Add(first);

            // Allow both same-type and mixed nested lists under the current item.
            ConsumeNestedBlocksForListItem(lines, ref j, level0Abs, firstContinuationIndent, options, state, first, allowNestedOrdered: true, allowNestedUnordered: true);

            while (j < lines.Length) {
                bool separatedByBlankLine = false;
                int itemStart = j;
                while (itemStart < lines.Length && string.IsNullOrWhiteSpace(lines[itemStart])) {
                    separatedByBlankLine = true;
                    itemStart++;
                }

                if (itemStart >= lines.Length
                    || LooksLikeHr(lines[itemStart])
                    || !IsUnorderedListLine(lines[itemStart], out var lvlAbs, out var isTask2, out var done2, out var content2)
                    || lvlAbs < level0Abs
                    || !TryGetUnorderedListMarkerInfo(lines[itemStart], out int markerIndent, out _, out char marker)
                    || (options.StrictListIndentation && markerIndent - firstMarkerIndent > 3)
                    || marker != firstMarker) {
                    break;
                }

                if (separatedByBlankLine) {
                    first.ForceLoose = true;
                }

                if (isTask2 && !options.TaskLists) {
                    isTask2 = false;
                    done2 = false;
                    content2 = GetUnorderedListItemContent(lines[itemStart]);
                }
                int continuationIndent = GetListContinuationIndent(lines[itemStart], options);
                int startColumn;
                if (!isTask2 && TryGetIndentedCodeListLead(lines[itemStart], options, out int itemCodeLeadIndent, out string itemCodeLeadContent, out int itemCodeLeadStartColumn)) {
                    continuationIndent = itemCodeLeadIndent;
                    content2 = itemCodeLeadContent;
                    startColumn = itemCodeLeadStartColumn;
                } else {
                    startColumn = GetListLeadContentStartColumn(lines[itemStart], options, isTask2);
                }
                int next = itemStart + 1;
                var itemSourceLines = new List<MarkdownSourceLineSlice>();
                var itemLines = ConsumeListContinuationLines(
                    lines,
                    ref next,
                    continuationIndent,
                    content2,
                    options,
                    breakOnAnyOrderedListLine: options.ListExtras,
                    sourceLines: itemSourceLines,
                    absoluteLineOffset: state.SourceLineOffset,
                    initialLineIndex: itemStart,
                    initialStartColumn: startColumn,
                    state: state);
                var li = CreateListItemFromLeadLines(itemLines, isTask2, done2, options, state, itemStart, itemSourceLines);
                li.Level = continuationIndentsByLevel != null
                    ? GetRelativeListItemLevel(continuationIndentsByLevel, lines[itemStart])
                    : lvlAbs - level0Abs;
                SetListItemMarkerSourceSpans(li, lines[itemStart], itemStart, isTask2, options, state);
                if (separatedByBlankLine) {
                    li.ForceLoose = true;
                }
                if (continuationIndentsByLevel != null) {
                    TrackListItemContinuationIndent(continuationIndentsByLevel, li.Level, continuationIndent);
                }
                AddListItemLeadSyntaxNodes(li, itemLines, itemStart, options, state, itemSourceLines);
                ul.Items.Add(li);
                j = next;

                ConsumeNestedBlocksForListItem(lines, ref j, lvlAbs, continuationIndent, options, state, li, allowNestedOrdered: true, allowNestedUnordered: true);
            }
            doc.Add(ul); i = j; return true;
        }
    }
}

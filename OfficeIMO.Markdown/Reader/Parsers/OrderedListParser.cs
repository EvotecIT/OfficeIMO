namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class OrderedListParser : IMarkdownBlockParser {
        private static bool TryStripTaskMarker(string? content, MarkdownReaderOptions options, out bool isTask, out bool done, out string stripped) {
            isTask = false;
            done = false;
            stripped = content ?? string.Empty;
            if (string.IsNullOrEmpty(stripped) || !options.TaskLists) return false;

            // Task marker is only valid at the start of the list item content: [ ] or [x] (case-insensitive).
            if (stripped.StartsWith("[ ]", StringComparison.Ordinal)) {
                isTask = true;
                done = false;
                stripped = stripped.Length > 4 && stripped[3] == ' ' ? stripped.Substring(4) : stripped.Substring(3);
                return true;
            }
            if (stripped.StartsWith("[x]", StringComparison.OrdinalIgnoreCase)) {
                isTask = true;
                done = true;
                stripped = stripped.Length > 4 && stripped[3] == ' ' ? stripped.Substring(4) : stripped.Substring(3);
                return true;
            }
            return false;
        }

        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.OrderedLists) return false;
            if (!IsOrderedListLine(lines[i], out int lvl0Abs, out int startNum, out var firstContent)) return false;
            if (!TryGetOrderedListMarkerInfo(lines[i], out _, out _, out _, out char firstDelimiter)) return false;
            var ol = new OrderedListBlock { Start = startNum };
            var continuationIndentsByLevel = options.StrictListIndentation ? new List<int>() : null;
            int firstContinuationIndent = GetListContinuationIndent(lines[i]);
            int firstStartColumn;

            int j = i + 1;
            bool firstIsTask = TryStripTaskMarker(firstContent, options, out _, out bool firstDone, out var strippedFirst);
            if (!firstIsTask && TryGetIndentedCodeListLead(lines[i], out int codeLeadIndent, out string codeLeadContent, out int codeLeadStartColumn)) {
                firstContinuationIndent = codeLeadIndent;
                strippedFirst = codeLeadContent;
                firstStartColumn = codeLeadStartColumn;
            } else {
                firstStartColumn = GetListLeadContentStartColumn(lines[i], firstIsTask);
            }
            var firstSourceLines = new List<MarkdownSourceLineSlice>();
            var firstLines = ConsumeListContinuationLines(
                lines,
                ref j,
                firstContinuationIndent,
                strippedFirst,
                options,
                breakOnAnyOrderedListLine: true,
                sourceLines: firstSourceLines,
                absoluteLineOffset: state.SourceLineOffset,
                initialLineIndex: i,
                initialStartColumn: firstStartColumn);
            var first = CreateListItemFromLeadLines(firstLines, firstIsTask, firstDone, options, state, firstSourceLines);
            first.Level = 0;
            if (continuationIndentsByLevel != null) {
                TrackListItemContinuationIndent(continuationIndentsByLevel, first.Level, firstContinuationIndent);
            }
            AddListItemLeadSyntaxNodes(first, firstLines, i, options, state, firstSourceLines);
            ol.Items.Add(first);

            ConsumeNestedBlocksForListItem(lines, ref j, lvl0Abs, firstContinuationIndent, options, state, first, allowNestedOrdered: true, allowNestedUnordered: true);

            while (j < lines.Length) {
                bool separatedByBlankLine = false;
                int itemStart = j;
                while (itemStart < lines.Length && string.IsNullOrWhiteSpace(lines[itemStart])) {
                    separatedByBlankLine = true;
                    itemStart++;
                }

                if (itemStart >= lines.Length
                    || !IsOrderedListLine(lines[itemStart], out var lvlAbs, out _, out var content)
                    || lvlAbs < lvl0Abs
                    || !TryGetOrderedListMarkerInfo(lines[itemStart], out _, out _, out _, out char delimiter)
                    || delimiter != firstDelimiter) {
                    break;
                }

                if (separatedByBlankLine) {
                    first.ForceLoose = true;
                }

                int continuationIndent = GetListContinuationIndent(lines[itemStart]);
                int next = itemStart + 1;
                bool isTask = TryStripTaskMarker(content, options, out _, out bool done, out var stripped);
                int startColumn;
                if (!isTask && TryGetIndentedCodeListLead(lines[itemStart], out int itemCodeLeadIndent, out string itemCodeLeadContent, out int itemCodeLeadStartColumn)) {
                    continuationIndent = itemCodeLeadIndent;
                    stripped = itemCodeLeadContent;
                    startColumn = itemCodeLeadStartColumn;
                } else {
                    startColumn = GetListLeadContentStartColumn(lines[itemStart], isTask);
                }
                var itemSourceLines = new List<MarkdownSourceLineSlice>();
                var itemLines = ConsumeListContinuationLines(
                    lines,
                    ref next,
                    continuationIndent,
                    stripped,
                    options,
                    breakOnAnyOrderedListLine: true,
                    sourceLines: itemSourceLines,
                    absoluteLineOffset: state.SourceLineOffset,
                    initialLineIndex: itemStart,
                    initialStartColumn: startColumn);
                var li = CreateListItemFromLeadLines(itemLines, isTask, done, options, state, itemSourceLines);
                li.Level = continuationIndentsByLevel != null
                    ? GetRelativeListItemLevel(continuationIndentsByLevel, lines[itemStart])
                    : lvlAbs - lvl0Abs;
                if (separatedByBlankLine) {
                    li.ForceLoose = true;
                }
                if (continuationIndentsByLevel != null) {
                    TrackListItemContinuationIndent(continuationIndentsByLevel, li.Level, continuationIndent);
                }
                AddListItemLeadSyntaxNodes(li, itemLines, itemStart, options, state, itemSourceLines);
                ol.Items.Add(li);
                j = next;

                ConsumeNestedBlocksForListItem(lines, ref j, lvlAbs, continuationIndent, options, state, li, allowNestedOrdered: true, allowNestedUnordered: true);
            }
            doc.Add(ol); i = j; return true;
        }
    }
}

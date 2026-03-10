namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class OrderedListParser : IMarkdownBlockParser {
        private static bool TryStripTaskMarker(string? content, out bool isTask, out bool done, out string stripped) {
            isTask = false;
            done = false;
            stripped = content ?? string.Empty;
            if (string.IsNullOrEmpty(stripped)) return false;

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
            var ol = new OrderedListBlock { Start = startNum };
            int firstContinuationIndent = GetListContinuationIndent(lines[i]);

            int j = i + 1;
            bool firstIsTask = TryStripTaskMarker(firstContent, out _, out bool firstDone, out var strippedFirst);
            var firstLines = ConsumeListContinuationLines(lines, ref j, firstContinuationIndent, strippedFirst, options, breakOnAnyOrderedListLine: true);
            var first = CreateListItemFromLeadLines(firstLines, firstIsTask, firstDone, options, state);
            first.Level = 0;
            AddListItemLeadSyntaxNodes(first, firstLines, i, options, state);
            ol.Items.Add(first);

            ConsumeNestedBlocksForListItem(lines, ref j, lvl0Abs, firstContinuationIndent, options, state, first, allowNestedOrdered: true, allowNestedUnordered: true);

            while (j < lines.Length && IsOrderedListLine(lines[j], out var lvlAbs, out _, out var content) && lvlAbs >= lvl0Abs) {
                int continuationIndent = GetListContinuationIndent(lines[j]);
                int next = j + 1;
                bool isTask = TryStripTaskMarker(content, out _, out bool done, out var stripped);
                var itemLines = ConsumeListContinuationLines(lines, ref next, continuationIndent, stripped, options, breakOnAnyOrderedListLine: true);
                var li = CreateListItemFromLeadLines(itemLines, isTask, done, options, state);
                li.Level = lvlAbs - lvl0Abs;
                AddListItemLeadSyntaxNodes(li, itemLines, j, options, state);
                ol.Items.Add(li);
                j = next;

                ConsumeNestedBlocksForListItem(lines, ref j, lvlAbs, continuationIndent, options, state, li, allowNestedOrdered: true, allowNestedUnordered: true);
            }
            doc.Add(ol); i = j; return true;
        }
    }
}

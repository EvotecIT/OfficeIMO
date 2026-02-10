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

            int j = i + 1;
            bool firstIsTask = TryStripTaskMarker(firstContent, out _, out bool firstDone, out var strippedFirst);
            var firstLines = ConsumeListContinuationLines(lines, ref j, lvl0Abs, strippedFirst, options);
            var firstParas = ParseParagraphsFromLines(firstLines, options, state);
            var first = firstIsTask ? ListItem.TaskInlines(firstParas[0], firstDone) : new ListItem(firstParas[0]);
            first.Level = 0;
            for (int p = 1; p < firstParas.Count; p++) first.AdditionalParagraphs.Add(firstParas[p]);
            ol.Items.Add(first);

            ConsumeNestedBlocksForListItem(lines, ref j, lvl0Abs, options, state, first, allowNestedOrdered: false, allowNestedUnordered: true);

            while (j < lines.Length && IsOrderedListLine(lines[j], out var lvlAbs, out _, out var content) && lvlAbs >= lvl0Abs) {
                int next = j + 1;
                bool isTask = TryStripTaskMarker(content, out _, out bool done, out var stripped);
                var itemLines = ConsumeListContinuationLines(lines, ref next, lvlAbs, stripped, options);
                var paras = ParseParagraphsFromLines(itemLines, options, state);
                var li = isTask ? ListItem.TaskInlines(paras[0], done) : new ListItem(paras[0]);
                li.Level = lvlAbs - lvl0Abs;
                for (int p = 1; p < paras.Count; p++) li.AdditionalParagraphs.Add(paras[p]);
                ol.Items.Add(li);
                j = next;

                ConsumeNestedBlocksForListItem(lines, ref j, lvlAbs, options, state, li, allowNestedOrdered: false, allowNestedUnordered: true);
            }
            doc.Add(ol); i = j; return true;
        }
    }
}

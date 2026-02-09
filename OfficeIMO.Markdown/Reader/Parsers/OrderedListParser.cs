namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class OrderedListParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.OrderedLists) return false;
            if (!IsOrderedListLine(lines[i], out int lvl0Abs, out int startNum, out var firstContent)) return false;
            var ol = new OrderedListBlock { Start = startNum };

            int j = i + 1;
            var firstLines = ConsumeListContinuationLines(lines, ref j, lvl0Abs, firstContent, options);
            var firstParas = ParseParagraphsFromLines(firstLines, options, state);
            var first = new ListItem(firstParas[0]) { Level = 0 };
            for (int p = 1; p < firstParas.Count; p++) first.AdditionalParagraphs.Add(firstParas[p]);
            ol.Items.Add(first);

            ConsumeNestedBlocksForListItem(lines, ref j, lvl0Abs, options, state, first, allowNestedOrdered: false, allowNestedUnordered: true);

            while (j < lines.Length && IsOrderedListLine(lines[j], out var lvlAbs, out _, out var content) && lvlAbs >= lvl0Abs) {
                int next = j + 1;
                var itemLines = ConsumeListContinuationLines(lines, ref next, lvlAbs, content, options);
                var paras = ParseParagraphsFromLines(itemLines, options, state);
                var li = new ListItem(paras[0]) { Level = lvlAbs - lvl0Abs };
                for (int p = 1; p < paras.Count; p++) li.AdditionalParagraphs.Add(paras[p]);
                ol.Items.Add(li);
                j = next;

                ConsumeNestedBlocksForListItem(lines, ref j, lvlAbs, options, state, li, allowNestedOrdered: false, allowNestedUnordered: true);
            }
            doc.Add(ol); i = j; return true;
        }
    }
}

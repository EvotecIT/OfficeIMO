namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class UnorderedListParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.UnorderedLists) return false;
            if (!IsUnorderedListLine(lines[i], out int level0Abs, out var isTask, out var done, out var firstContent)) return false;
            var ul = new UnorderedListBlock();

            int j = i + 1;
            var firstLines = ConsumeListContinuationLines(lines, ref j, level0Abs, firstContent, options);
            var firstParas = ParseParagraphsFromLines(firstLines, options, state);
            var firstInline = firstParas[0];
            var first = isTask ? ListItem.TaskInlines(firstInline, done) : new ListItem(firstInline);
            for (int p = 1; p < firstParas.Count; p++) first.AdditionalParagraphs.Add(firstParas[p]);
            first.Level = 0;
            ul.Items.Add(first);

            // Mixed nesting: allow an indented ordered list or fenced code block to be attached to the current item.
            ConsumeNestedBlocksForListItem(lines, ref j, level0Abs, options, state, first, allowNestedOrdered: true, allowNestedUnordered: false);

            while (j < lines.Length && IsUnorderedListLine(lines[j], out var lvlAbs, out var isTask2, out var done2, out var content2) && lvlAbs >= level0Abs) {
                int next = j + 1;
                var itemLines = ConsumeListContinuationLines(lines, ref next, lvlAbs, content2, options);
                var paras = ParseParagraphsFromLines(itemLines, options, state);
                var inline = paras[0];
                var li = isTask2 ? ListItem.TaskInlines(inline, done2) : new ListItem(inline);
                for (int p = 1; p < paras.Count; p++) li.AdditionalParagraphs.Add(paras[p]);
                li.Level = lvlAbs - level0Abs;
                ul.Items.Add(li);
                j = next;

                ConsumeNestedBlocksForListItem(lines, ref j, lvlAbs, options, state, li, allowNestedOrdered: true, allowNestedUnordered: false);
            }
            doc.Add(ul); i = j; return true;
        }
    }
}

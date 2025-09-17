namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class UnorderedListParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.UnorderedLists) return false;
            if (!IsUnorderedListLine(lines[i], out int level0, out var isTask, out var done, out var firstContent)) return false;
            var ul = new UnorderedListBlock();
            var first = isTask ? ListItem.Task(firstContent, done) : ListItem.Text(firstContent);
            first.Level = level0;
            ul.Items.Add(first);
            int j = i + 1;
            while (j < lines.Length && IsUnorderedListLine(lines[j], out var lvl, out var isTask2, out var done2, out var content2)) {
                var li = isTask2 ? ListItem.Task(content2, done2) : ListItem.Text(content2);
                li.Level = lvl;
                ul.Items.Add(li);
                j++;
            }
            doc.Add(ul); i = j; return true;
        }
    }
}

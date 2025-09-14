namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class UnorderedListParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.UnorderedLists) return false;
            if (!IsUnorderedListLine(lines[i], out var isTask, out var done, out var firstContent)) return false;
            var ul = new UnorderedListBlock();
            if (isTask) ul.Items.Add(ListItem.Task(firstContent, done)); else ul.Items.Add(ListItem.Text(firstContent));
            int j = i + 1;
            while (j < lines.Length && IsUnorderedListLine(lines[j], out var isTask2, out var done2, out var content2)) {
                if (isTask2) ul.Items.Add(ListItem.Task(content2, done2)); else ul.Items.Add(ListItem.Text(content2));
                j++;
            }
            doc.Add(ul); i = j; return true;
        }
    }
}

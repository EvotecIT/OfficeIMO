namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class OrderedListParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.OrderedLists) return false;
            if (!IsOrderedListLine(lines[i], out int startNum, out var firstContent)) return false;
            var ol = new OrderedListBlock { Start = startNum };
            ol.Items.Add(new ListItem(ParseInlines(firstContent)));
            int j = i + 1;
            while (j < lines.Length && IsOrderedListLine(lines[j], out _, out var content)) { ol.Items.Add(new ListItem(ParseInlines(content))); j++; }
            doc.Add(ol); i = j; return true;
        }
    }
}

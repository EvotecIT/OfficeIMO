namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class TableParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.Tables) return false;
            if (!LooksLikeTableRow(lines[i])) return false;
            int start = i; int j = i;
            while (j < lines.Length && LooksLikeTableRow(lines[j])) j++;
            var table = ParseTable(lines, start, j - 1);
            doc.Add(table); i = j; return true;
        }
    }
}

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class TableParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.Tables) return false;
            if (!LooksLikeTableRow(lines[i])) return false;
            if (!TryGetTableExtent(lines, i, out int end, out _)) return false;

            var table = ParseTable(lines, i, end);
            doc.Add(table); i = end + 1; return true;
        }
    }
}

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class TableParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.Tables) return false;
            if (!LooksLikeTableRow(lines[i])) return false;
            bool hasOuterPipes = false;
            var firstTrimmed = lines[i].Trim();
            if (firstTrimmed.Length > 0 && (firstTrimmed[0] == '|' || firstTrimmed[firstTrimmed.Length - 1] == '|')) {
                hasOuterPipes = true;
            }

            int start = i; int j = i;
            while (j < lines.Length && LooksLikeTableRow(lines[j])) j++;

            if (!hasOuterPipes && j == start + 1) return false;

            var table = ParseTable(lines, start, j - 1);
            doc.Add(table); i = j; return true;
        }
    }
}

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class FrontMatterParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.FrontMatter) return false;
            if (i != 0) return false; // only at very top
            if (i < lines.Length && lines[i].Trim() == "---") {
                int start = i + 1; int end = -1;
                for (int j = start; j < lines.Length; j++) { if (lines[j].Trim() == "---") { end = j; break; } }
                if (end > start) {
                    var dict = ParseFrontMatter(lines, start, end - 1);
                    if (dict.Count > 0) doc.Add(FrontMatterBlock.FromObject(dict));
                    i = end + 1;
                    if (i < lines.Length && string.IsNullOrWhiteSpace(lines[i])) i++;
                    return true;
                }
            }
            return false;
        }
    }
}

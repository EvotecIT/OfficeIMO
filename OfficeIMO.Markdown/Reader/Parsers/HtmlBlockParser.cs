namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class HtmlBlockParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            // Simple heuristic: a line starting with '<' and not a known inline-only tag
            var t = lines[i].TrimStart();
            if (t.Length == 0 || t[0] != '<') return false;
            // Avoid consuming plain inline <u> ... handled by inline parser inside paragraphs
            if (t.StartsWith("<u>") && t.EndsWith("</u>")) return false;
            int j = i;
            var sb = new StringBuilder();
            while (j < lines.Length && !string.IsNullOrWhiteSpace(lines[j])) { sb.AppendLine(lines[j]); j++; }
            doc.Add(new HtmlRawBlock(sb.ToString().TrimEnd()));
            i = j; return true;
        }
    }
}

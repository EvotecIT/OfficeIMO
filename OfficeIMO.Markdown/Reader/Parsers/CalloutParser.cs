namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class CalloutParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.Callouts) return false;
            if (!IsCalloutHeader(lines[i], out string kind, out string title)) return false;
            var body = new System.Text.StringBuilder();
            int j = i + 1;
            while (j < lines.Length && lines[j].StartsWith("> ")) { body.AppendLine(lines[j].Substring(2)); j++; }
            doc.Add(new CalloutBlock(kind, title, body.ToString().TrimEnd()));
            i = j; return true;
        }
    }
}

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class HeadingParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.Headings) return false;
            if (!IsAtxHeading(lines[i], out int level, out string text)) return false;
            doc.Add(new HeadingBlock(level, text));
            i++; return true;
        }
    }
}

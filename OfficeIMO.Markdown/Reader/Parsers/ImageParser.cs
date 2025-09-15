namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class ImageParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.Images) return false;
            if (!TryParseImage(lines[i], out var img)) return false;
            int j = i + 1;
            if (j < lines.Length && TryParseCaption(lines[j], out var cap)) { img.Caption = cap; j++; }
            doc.Add(img); i = j; return true;
        }
    }
}

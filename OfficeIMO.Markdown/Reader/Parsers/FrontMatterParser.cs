namespace OfficeIMO.Markdown;

internal static partial class MarkdownReader {
    internal sealed class FrontMatterParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownLoadOptions options, MarkdownDocument doc, MarkdownReaderState state) {
            if (!options.FrontMatter || i != 0 || !TryLocateFrontMatter(lines, out int closingIndex)) return false;
            doc.Add(ParseFrontMatterBlock(lines, closingIndex, state));
            i = closingIndex + 1;
            if (i < lines.Length && string.IsNullOrWhiteSpace(lines[i])) i++;
            return true;
        }
    }
}

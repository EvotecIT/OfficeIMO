namespace OfficeIMO.Markdown;

internal static partial class MarkdownReader {
    internal sealed class FrontMatterParser : MarkdownSourceLinesBlockParser {
        public override bool TryParse(MarkdownSourceLines lines, ref int i, MarkdownLoadOptions options, MarkdownDocument doc, MarkdownReaderState state) {
            if (!options.FrontMatter || i != 0 || !TryLocateFrontMatter(lines, out int closingIndex)) return false;
            doc.Add(ParseFrontMatterBlock(lines, closingIndex, state));
            i = closingIndex + 1;
            if (i < lines.Length && lines.IsBlank(i)) i++;
            return true;
        }
    }
}

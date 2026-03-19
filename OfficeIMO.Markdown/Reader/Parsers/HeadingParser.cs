namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class HeadingParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.Headings) return false;
            if (!TryGetAtxHeadingContentRange(lines[i], out int level, out int contentStart, out int contentEnd, out string text)) return false;
            var sourceMap = BuildInlineSourceMapForSingleLine(text, state.SourceLineOffset + i + 1, contentStart + 1, state);
            doc.Add(new HeadingBlock(level, ParseInlines(text, options, state, sourceMap)));
            i++; return true;
        }
    }
}

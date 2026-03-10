namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    /// <summary>
    /// Parses reference-style link definitions: [label]: url "title".
    /// Definitions are stored in state and the lines are consumed (not added to the doc).
    /// </summary>
    internal sealed class ReferenceLinkDefParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (TryParseReferenceLinkDefinition(lines, i, options, out var label, out var url, out var title, out var consumedLines)) {
                var resolved = ResolveUrl(url, options);
                if (resolved != null) {
                    state.LinkRefs[label] = (resolved!, title);
                }
                i += consumedLines;
                return true;
            }
            return false;
        }
    }
}

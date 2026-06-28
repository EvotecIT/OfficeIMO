namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    /// <summary>
    /// Parses reference-style link definitions: [label]: url "title".
    /// Definitions are stored in state and the lines are consumed (not added to the doc).
    /// </summary>
    internal sealed class ReferenceLinkDefParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (TryParseReferenceLinkDefinition(
                lines,
                i,
                options,
                state,
                out var label,
                out var url,
                out var title,
                out var consumedLines,
                out var labelSpan,
                out var urlSpan,
                out var titleSpan,
                out var openingMarkerSpan,
                out var separatorMarkerSpan)) {
                var resolved = ResolveUrl(url, options);
                if (resolved != null && !state.LinkRefs.ContainsKey(label)) {
                    var sourceSpan = CreateLineSpan(
                        state,
                        state.SourceLineOffset + i + 1,
                        state.SourceLineOffset + i + consumedLines);
                    state.LinkRefs[label] = new MarkdownReferenceLinkDefinition(
                        label,
                        resolved!,
                        title,
                        sourceSpan,
                        labelSpan,
                        urlSpan,
                        titleSpan,
                        openingMarkerSpan,
                        separatorMarkerSpan);
                }
                i += consumedLines;
                return true;
            }
            return false;
        }
    }
}

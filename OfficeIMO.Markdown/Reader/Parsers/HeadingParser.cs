namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class HeadingParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.Headings) return false;
            if (!TryGetAtxHeadingContentRange(lines[i], out int level, out int contentStart, out int contentEnd, out string text, out int closingMarkerStart, out int closingMarkerEnd)) return false;
            var sourceMap = BuildInlineSourceMapForSingleLine(text, state.SourceLineOffset + i + 1, contentStart + 1, state);
            var heading = new HeadingBlock(level, ParseInlines(text, options, state, sourceMap));
            var markerStartColumn = CountLeadingSpaces(lines[i]) + 1;
            heading.SetLevelSourceInfo(0, markerStartColumn, markerStartColumn + level - 1);
            if (contentEnd > contentStart) {
                heading.SetTextSourceInfo(0, contentStart + 1, contentEnd);
            }
            if (closingMarkerStart >= 0 && closingMarkerEnd > closingMarkerStart) {
                var absoluteLineNumber = state.SourceLineOffset + i + 1;
                heading.SetClosingMarkerSourceInfo(
                    0,
                    closingMarkerStart + 1,
                    closingMarkerEnd,
                    CreateSpan(state, absoluteLineNumber, closingMarkerStart + 1, absoluteLineNumber, closingMarkerEnd));
            }
            doc.Add(heading);
            i++; return true;
        }
    }
}

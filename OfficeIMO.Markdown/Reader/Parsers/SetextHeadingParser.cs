namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    /// <summary>
    /// Parses Setext headings:
    ///   Title
    ///   =====  (level 1)
    ///   Title
    ///   -----  (level 2)
    /// Requires one or more underline characters and no other content on the underline line.
    /// </summary>
    internal sealed class SetextHeadingParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.Headings) return false;
            if (i + 1 >= lines.Length) return false;
            var line = lines[i];
            var next = lines[i + 1];
            if (string.IsNullOrWhiteSpace(line) || string.IsNullOrWhiteSpace(next)) return false;
            if (!TryGetSetextHeadingUnderlineLevel(next, out int level)) return false;
            var t = next.Trim();
            var headingText = line.Trim();
            var contentStart = line.IndexOf(headingText, StringComparison.Ordinal);
            var sourceMap = BuildInlineSourceMapForSingleLine(headingText, state.SourceLineOffset + i + 1, contentStart + 1, state);
            var heading = new HeadingBlock(level, ParseInlines(headingText, options, state, sourceMap));
            var markerStartColumn = next.IndexOf(t, StringComparison.Ordinal) + 1;
            heading.SetLevelSourceInfo(1, markerStartColumn, markerStartColumn + t.Length - 1);
            if (headingText.Length > 0) {
                heading.SetTextSourceInfo(0, contentStart + 1, contentStart + headingText.Length);
            }
            doc.Add(heading);
            i += 2; // consume both lines
            return true;
        }
    }
}

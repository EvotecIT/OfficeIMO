namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    /// <summary>
    /// Parses Setext headings:
    ///   Title
    ///   =====  (level 1)
    ///   Title
    ///   -----  (level 2)
    /// Requires at least 3 underline characters and no other content on the underline line.
    /// </summary>
    internal sealed class SetextHeadingParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.Headings) return false;
            if (i + 1 >= lines.Length) return false;
            var line = lines[i];
            var next = lines[i + 1];
            if (string.IsNullOrWhiteSpace(line) || string.IsNullOrWhiteSpace(next)) return false;
            var t = next.Trim();
            // must be only '=' or '-' with length >= 3
            char ch = '\0';
            foreach (var c in t) {
                if (c == '=' || c == '-') { if (ch == '\0') ch = c; if (c != ch) return false; continue; }
                return false;
            }
            if (t.Length < 3) return false;
            int level = ch == '=' ? 1 : 2;
            doc.Add(new HeadingBlock(level, line.Trim()));
            i += 2; // consume both lines
            return true;
        }
    }
}


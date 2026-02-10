namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    /// <summary>
    /// Parses reference-style link definitions: [label]: url "title".
    /// Definitions are stored in state and the lines are consumed (not added to the doc).
    /// </summary>
    internal sealed class ReferenceLinkDefParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            var line = lines[i];
            if (string.IsNullOrWhiteSpace(line)) return false;

            // Do not treat indented code as reference definitions.
            int leading = 0;
            while (leading < line.Length && line[leading] == ' ') leading++;
            if (leading >= 4) return false;
            if (leading < line.Length && line[leading] == '\t') return false;

            var t = line.Trim();
            if (t.Length < 5 || t[0] != '[') return false;
            if (t.Length > 1 && t[1] == '^') return false; // footnote definition, not a link ref
            int rb = t.IndexOf(']');
            if (rb <= 1) return false;
            if (rb + 1 >= t.Length || t[rb + 1] != ':') return false;
            string label = NormalizeReferenceLabel(t.Substring(1, rb - 1));
            string rest = t.Substring(rb + 2).Trim();
            if (string.IsNullOrEmpty(rest)) return false;
            if (!TrySplitUrlAndOptionalTitle(rest, out var url, out var title)) return false;
            if (!string.IsNullOrEmpty(label) && !string.IsNullOrEmpty(url))
            {
                var resolved = ResolveUrl(url, options);
                if (!string.IsNullOrEmpty(resolved)) {
                    state.LinkRefs[label] = (resolved!, title);
                }
            }
            i++; return true;
        }
    }
}

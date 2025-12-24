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
            var t = line.Trim();
            if (t.Length < 5 || t[0] != '[') return false;
            if (t.Length > 1 && t[1] == '^') return false; // footnote definition, not a link ref
            int rb = t.IndexOf(']');
            if (rb <= 1) return false;
            if (rb + 1 >= t.Length || t[rb + 1] != ':') return false;
            string label = t.Substring(1, rb - 1);
            string rest = t.Substring(rb + 2).Trim();
            if (string.IsNullOrEmpty(rest)) return false;
            // url (optionally in < >) + optional "title"
            string url = rest; string? title = null;
            if (rest[0] == '<') {
                int gt = rest.IndexOf('>'); if (gt > 1) { url = rest.Substring(1, gt - 1); rest = rest.Substring(gt + 1).Trim(); }
            }
            int q = rest.IndexOf('"');
            if (q >= 0) {
                url = rest.Substring(0, q).Trim();
                int q2 = rest.LastIndexOf('"');
                if (q2 > q) title = rest.Substring(q + 1, q2 - q - 1);
            }
            if (!string.IsNullOrEmpty(label) && !string.IsNullOrEmpty(url))
            {
                var resolved = ResolveUrl(url, options) ?? url;
                state.LinkRefs[label] = (resolved, title);
            }
            i++; return true;
        }
    }
}

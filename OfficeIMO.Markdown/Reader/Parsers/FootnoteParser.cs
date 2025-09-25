namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class FootnoteParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            var line = lines[i];
            if (string.IsNullOrWhiteSpace(line)) return false;
            var t = line.TrimStart();
            if (!(t.Length > 4 && t[0] == '[' && t.Length > 2 && t[1] == '^')) return false;
            int rb = t.IndexOf(']'); if (rb < 0) return false;
            if (rb + 1 >= t.Length || t[rb + 1] != ':') return false;
            string label = t.Substring(2, rb - 2);
            string content = t.Substring(rb + 2).TrimStart();
            // Continuation lines: indented by at least two spaces or a tab
            int j = i + 1;
            while (j < lines.Length) {
                var ln = lines[j];
                if (string.IsNullOrEmpty(ln)) break;
                int leading = 0; while (leading < ln.Length && ln[leading] == ' ') leading++;
                if (leading >= 2 || (leading < ln.Length && ln[leading] == '\t')) { content += "\n" + ln.TrimStart(); j++; } else break;
            }
            doc.Add(new FootnoteDefinitionBlock(label, content));
            i = j; return true;
        }
    }
}


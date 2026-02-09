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

            // Collect content lines; allow blank lines inside the footnote only when followed by indented continuation.
            var contentLines = new List<string> { t.Substring(rb + 2).TrimStart() };

            int j = i + 1;
            while (j < lines.Length) {
                var ln = lines[j] ?? string.Empty;

                if (ln.Length == 0) {
                    int peek = j + 1;
                    if (peek >= lines.Length) break;
                    var next = lines[peek] ?? string.Empty;
                    int leadingNext = 0; while (leadingNext < next.Length && next[leadingNext] == ' ') leadingNext++;
                    bool indentedNext = leadingNext >= 2 || (leadingNext < next.Length && next[leadingNext] == '\t');
                    if (!indentedNext) break;

                    contentLines.Add(string.Empty); // paragraph separator
                    j++;
                    continue;
                }

                int leading = 0; while (leading < ln.Length && ln[leading] == ' ') leading++;
                if (leading >= 2 || (leading < ln.Length && ln[leading] == '\t')) {
                    contentLines.Add(ln.TrimStart());
                    j++;
                    continue;
                }

                break;
            }

            string content = string.Join("\n", contentLines);
            var paragraphs = ParseParagraphsFromLines(contentLines, options, state);

            doc.Add(new FootnoteDefinitionBlock(label, content, paragraphs));
            i = j; return true;
        }
    }
}


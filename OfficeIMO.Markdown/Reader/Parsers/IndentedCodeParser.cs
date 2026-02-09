using System.Text;

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    /// <summary>
    /// Parses indented code blocks (4-space indented). This is a pragmatic subset intended to improve
    /// compatibility with common Markdown sources; output is represented as a fenced <see cref="CodeBlock"/>.
    /// </summary>
    internal sealed class IndentedCodeParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.IndentedCodeBlocks) return false;
            if (lines == null || i < 0 || i >= lines.Length) return false;

            var line = lines[i] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(line)) return false;

            int indent = CountLeadingSpaces(line);
            if (indent < 4) return false;

            var sb = new StringBuilder();
            int j = i;

            while (j < lines.Length) {
                var cur = lines[j] ?? string.Empty;

                if (string.IsNullOrWhiteSpace(cur)) {
                    // Include blank lines only if there is a following indented line (otherwise end block).
                    int peek = j + 1;
                    if (peek >= lines.Length) break;
                    int nextIndent = CountLeadingSpaces(lines[peek] ?? string.Empty);
                    if (nextIndent < 4) break;
                    sb.AppendLine();
                    j++;
                    continue;
                }

                int curIndent = CountLeadingSpaces(cur);
                if (curIndent < 4) break;

                // Strip the first 4 spaces; preserve any additional indentation.
                sb.AppendLine(cur.Substring(4));
                j++;
            }

            // Keep content as-is (minus the last newline we appended via AppendLine).
            string content = sb.ToString().TrimEnd('\r', '\n');
            doc.Add(new CodeBlock(string.Empty, content));
            i = j;
            return true;
        }
    }
}


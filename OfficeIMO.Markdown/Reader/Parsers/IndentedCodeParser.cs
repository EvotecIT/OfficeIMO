using System.Text;

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    /// <summary>
    /// Parses indented code blocks (4-space indented). This is a pragmatic subset intended to improve
    /// compatibility with common Markdown sources; output is represented as a fenced <see cref="CodeBlock"/>.
    /// </summary>
    internal sealed class IndentedCodeParser : IMarkdownBlockParser {
        private const int IndentedCodeMinimumSpaces = 4;

        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.IndentedCodeBlocks) return false;
            if (lines == null || i < 0 || i >= lines.Length) return false;

            var line = lines[i] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(line)) return false;

            int indent = CountLeadingIndentColumns(line);
            if (indent < IndentedCodeMinimumSpaces) return false;

            var sb = new StringBuilder();
            int j = i;

            while (j < lines.Length) {
                var cur = lines[j] ?? string.Empty;

                if (string.IsNullOrWhiteSpace(cur)) {
                    // Include blank lines only if there is a following indented line (otherwise end block).
                    if (!HasIndentedCodeContinuationAfterBlankLines(lines, j, IndentedCodeMinimumSpaces)) break;
                    sb.AppendLine();
                    j++;
                    continue;
                }

                int curIndent = CountLeadingIndentColumns(cur);
                if (curIndent < IndentedCodeMinimumSpaces) break;

                // Strip the first four indentation columns; preserve any additional indentation.
                sb.AppendLine(StripLeadingIndentColumns(cur, IndentedCodeMinimumSpaces));
                j++;
            }

            // Keep content as-is (minus the last newline we appended via AppendLine).
            string content = RemoveSingleTrailingLineEnding(sb.ToString());
            if (TryTakePendingGenericAttributeBlock(state, out var pending)) {
                var paragraphText = JoinParagraphLines(new List<string>(content.Split('\n')), options).Trim();
                var paragraph = new ParagraphBlock(ParseInlines(paragraphText, options, state));
                paragraph.SetAttributes(pending.Attributes);
                MarkdownGenericAttributeSourceSpans.Set(paragraph, pending.SourceText, pending.SourceSpan);
                doc.Add(paragraph);
                i = j;
                return true;
            }

            doc.Add(new CodeBlock(string.Empty, content, isFenced: false));
            i = j;
            return true;
        }
    }
}

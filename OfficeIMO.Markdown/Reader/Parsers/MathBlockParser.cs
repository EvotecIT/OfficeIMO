namespace OfficeIMO.Markdown;

internal static partial class MarkdownReader {
    /// <summary>
    /// Parses display math when <see cref="MarkdownLoadOptions.Math"/> is enabled: either a line holding only
    /// <c>$$...$$</c>, or an opening <c>$$</c> line, TeX lines and a closing <c>$$</c> line. An unclosed block is
    /// left to the paragraph parser.
    /// </summary>
    internal sealed class MathBlockParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownLoadOptions options, MarkdownDocument doc, MarkdownReaderState state) {
            if (!options.Math || lines == null || i < 0 || i >= lines.Length) return false;
            string line = lines[i] ?? string.Empty;
            int indent = CountMathIndent(line);
            if (indent > 3 || string.CompareOrdinal(line, indent, "$$", 0, 2) != 0) return false;
            string trimmed = line.Substring(indent).TrimEnd();

            if (trimmed.Length > 4 && trimmed.EndsWith("$$", StringComparison.Ordinal)) {
                string inner = trimmed.Substring(2, trimmed.Length - 4);
                if (inner.Trim().Length == 0 || inner.IndexOf("$$", StringComparison.Ordinal) >= 0) return false;
                doc.Add(new MathBlock(inner) { IsSingleLine = true, IndentColumns = indent });
                i++;
                return true;
            }

            if (trimmed.Length != 2) return false;
            int lineLimit = lines.Length - (state.HasSyntheticTerminalLine ? 1 : 0);
            int j = i + 1;
            while (j < lineLimit && !IsMathClosingLine(lines[j])) {
                if (string.IsNullOrWhiteSpace(lines[j])) return false;
                j++;
            }

            if (j >= lineLimit || j == i + 1) return false;
            var content = new System.Text.StringBuilder();
            for (int k = i + 1; k < j; k++) {
                if (k > i + 1) content.Append('\n');
                content.Append(StripLeadingIndentColumns(lines[k], indent));
            }

            doc.Add(new MathBlock(content.ToString()) { IndentColumns = indent });
            i = j + 1;
            return true;
        }

        private static bool IsMathClosingLine(string? line) {
            if (line == null) return false;
            int indent = CountMathIndent(line);
            return indent <= 3 && line.Substring(indent).TrimEnd() == "$$";
        }

        private static int CountMathIndent(string line) {
            int count = 0;
            while (count < line.Length && line[count] == ' ') count++;
            return count;
        }
    }
}

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class FootnoteParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            var line = lines[i];
            if (string.IsNullOrWhiteSpace(line)) return false;
            // Do not treat indented code as a footnote definition.
            int leading0 = 0; while (leading0 < line.Length && line[leading0] == ' ') leading0++;
            if (leading0 >= 4) return false;
            if (leading0 < line.Length && line[leading0] == '\t') return false;
            var t = line.TrimStart();
            if (!(t.Length > 4 && t[0] == '[' && t.Length > 2 && t[1] == '^')) return false;
            int rb = t.IndexOf(']'); if (rb < 0) return false;
            if (rb + 1 >= t.Length || t[rb + 1] != ':') return false;
            string label = t.Substring(2, rb - 2);

            int firstContentIndex = rb + 2;
            while (firstContentIndex < t.Length && char.IsWhiteSpace(t[firstContentIndex])) {
                firstContentIndex++;
            }

            int absoluteLine = state.SourceLineOffset + i + 1;
            var contentSourceLines = new List<MarkdownSourceLineSlice> {
                new MarkdownSourceLineSlice(
                    firstContentIndex < t.Length ? t.Substring(firstContentIndex) : string.Empty,
                    absoluteLine,
                    leading0 + firstContentIndex + 1)
            };

            int j = i + 1;
            while (j < lines.Length) {
                var ln = lines[j] ?? string.Empty;

                if (ln.Length == 0) {
                    int peek = j + 1;
                    if (peek >= lines.Length) break;
                    var next = lines[peek] ?? string.Empty;
                    bool indentedNext = CountLeadingIndentColumns(next) >= 2;
                    if (!indentedNext) break;

                    contentSourceLines.Add(new MarkdownSourceLineSlice(string.Empty, state.SourceLineOffset + j + 1, 1));
                    j++;
                    continue;
                }

                if (CountLeadingIndentColumns(ln) >= 2) {
                    contentSourceLines.Add(new MarkdownSourceLineSlice(
                        StripLeadingIndentColumns(ln, 2),
                        state.SourceLineOffset + j + 1,
                        GetStartColumnAfterStrippingIndent(ln, 2)));
                    j++;
                    continue;
                }

                break;
            }

            string content = string.Join("\n", contentSourceLines.Select(slice => slice.Text ?? string.Empty));
            var (blocks, syntaxChildren) = ParseFootnoteBody(contentSourceLines, options, state);

            doc.Add(new FootnoteDefinitionBlock(label, content, blocks, syntaxChildren));
            i = j; return true;
        }
    }

    private static (IReadOnlyList<IMarkdownBlock> Blocks, IReadOnlyList<MarkdownSyntaxNode> SyntaxChildren) ParseFootnoteBody(
        List<MarkdownSourceLineSlice> contentSourceLines,
        MarkdownReaderOptions options,
        MarkdownReaderState state) {
        if (contentSourceLines == null || contentSourceLines.Count == 0) {
            return (Array.Empty<IMarkdownBlock>(), Array.Empty<MarkdownSyntaxNode>());
        }

        var (blocks, syntaxChildren) = ParseNestedMarkdownBlocks(contentSourceLines, options, state);
        if (blocks.Count > 0) {
            return (blocks, syntaxChildren);
        }

        var paragraphs = ParseParagraphBlocksFromSourceLines(contentSourceLines, options, state);
        var paragraphSyntax = new List<MarkdownSyntaxNode>();
        AddParagraphSyntaxNodes(paragraphSyntax, contentSourceLines, options, state);
        return (paragraphs, paragraphSyntax);
    }
}


namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class DefinitionListParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.DefinitionLists) return false;
            if (!ShouldTreatAsDefinitionLine(lines, i, options)) return false;
            var dl = new DefinitionListBlock();
            dl.SetParsingContext(options, state);
            int j = i;
            while (j < lines.Length && ShouldTreatAsDefinitionLine(lines, j, options)) {
                if (!TryGetDefinitionSeparator(lines[j], out var idx)) break;
                var term = lines[j].Substring(0, idx).Trim();
                var def = lines[j].Substring(idx + 1).TrimStart();
                dl.Items.Add((term, def));

                int lineNumber = state.SourceLineOffset + j + 1;
                var lineSpan = new MarkdownSourceSpan(lineNumber, lineNumber);
                dl.SyntaxItems.Add(new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.DefinitionItem,
                    lineSpan,
                    term,
                    new[] {
                        new MarkdownSyntaxNode(MarkdownSyntaxKind.DefinitionTerm, lineSpan, term),
                        new MarkdownSyntaxNode(
                            MarkdownSyntaxKind.DefinitionValue,
                            lineSpan,
                            def,
                            new[] {
                                new MarkdownSyntaxNode(MarkdownSyntaxKind.Paragraph, lineSpan, def)
                            })
                    }));
                j++;
            }
            doc.Add(dl); i = j; return true;
        }
    }
}

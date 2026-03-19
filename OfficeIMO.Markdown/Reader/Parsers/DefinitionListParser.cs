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
                var line = lines[j] ?? string.Empty;
                int lineNumber = state.SourceLineOffset + j + 1;

                int termStartIndex = 0;
                while (termStartIndex < idx && char.IsWhiteSpace(line[termStartIndex])) {
                    termStartIndex++;
                }

                int termEndExclusive = idx;
                while (termEndExclusive > termStartIndex && char.IsWhiteSpace(line[termEndExclusive - 1])) {
                    termEndExclusive--;
                }

                int definitionStartIndex = idx + 1;
                while (definitionStartIndex < line.Length && char.IsWhiteSpace(line[definitionStartIndex])) {
                    definitionStartIndex++;
                }

                int definitionEndExclusive = line.Length;
                while (definitionEndExclusive > definitionStartIndex && char.IsWhiteSpace(line[definitionEndExclusive - 1])) {
                    definitionEndExclusive--;
                }

                var term = termStartIndex < termEndExclusive
                    ? line.Substring(termStartIndex, termEndExclusive - termStartIndex)
                    : string.Empty;
                var def = definitionStartIndex < definitionEndExclusive
                    ? line.Substring(definitionStartIndex, definitionEndExclusive - definitionStartIndex)
                    : string.Empty;

                var termSourceMap = BuildInlineSourceMapForSingleLine(term, lineNumber, termStartIndex + 1, state);
                var definitionSourceMap = BuildInlineSourceMapForSingleLine(def, lineNumber, definitionStartIndex + 1, state);
                var termInlines = ParseInlines(term, options, state, termSourceMap);
                var definitionInlines = ParseInlines(def, options, state, definitionSourceMap);
                var termSpan = CreateSpan(
                    state,
                    lineNumber,
                    termStartIndex + 1,
                    lineNumber,
                    Math.Max(termStartIndex + 1, termEndExclusive));
                var definitionSpan = CreateSpan(
                    state,
                    lineNumber,
                    definitionStartIndex + 1,
                    lineNumber,
                    Math.Max(definitionStartIndex + 1, definitionEndExclusive));
                var definitionSourceLines = new List<MarkdownSourceLineSlice> {
                    new MarkdownSourceLineSlice(def, lineNumber, definitionStartIndex + 1)
                };
                int next = j + 1;
                ConsumeDefinitionContinuationLines(lines, ref next, CountLeadingIndentColumns(line) + 2, state.SourceLineOffset, definitionSourceLines);
                var (definitionBlocks, definitionSyntaxChildren) = ParseDefinitionBody(definitionSourceLines, options, state);
                if (definitionSyntaxChildren.Count > 0) {
                    definitionSpan = MarkdownBlockSyntaxBuilder.GetAggregateSpan(definitionSyntaxChildren) ?? definitionSpan;
                }

                var termNode = MarkdownBlockSyntaxBuilder.BuildInlineContainerNode(
                    MarkdownSyntaxKind.DefinitionTerm,
                    termInlines,
                    termSpan,
                    term);
                var valueNode = new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.DefinitionValue,
                    definitionSpan,
                    RenderDefinitionLiteral(definitionBlocks, def),
                    definitionSyntaxChildren);
                dl.AddParsedGroup(
                    new DefinitionListGroup(
                        new[] { termInlines },
                        new[] { new DefinitionListDefinition(definitionBlocks) }),
                    new MarkdownSyntaxNode(
                        MarkdownSyntaxKind.DefinitionGroup,
                        MarkdownBlockSyntaxBuilder.GetAggregateSpan(new[] { termNode, valueNode }),
                        children: new[] {
                            termNode,
                            valueNode
                        }));
                j = next;
                int separator = j;
                while (separator < lines.Length && string.IsNullOrWhiteSpace(lines[separator])) {
                    separator++;
                }

                if (separator < lines.Length && ShouldTreatAsDefinitionLine(lines, separator, options)) {
                    j = separator;
                }
            }
            doc.Add(dl); i = j; return true;
        }
    }

    private static void ConsumeDefinitionContinuationLines(
        string[] lines,
        ref int index,
        int continuationIndent,
        int absoluteLineOffset,
        List<MarkdownSourceLineSlice> definitionSourceLines) {
        if (lines == null || definitionSourceLines == null) {
            return;
        }

        while (index < lines.Length) {
            var line = lines[index] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(line)) {
                int peek = index;
                while (peek < lines.Length && string.IsNullOrWhiteSpace(lines[peek])) {
                    peek++;
                }

                if (peek >= lines.Length || CountLeadingIndentColumns(lines[peek] ?? string.Empty) < continuationIndent) {
                    return;
                }

                while (index < peek) {
                    definitionSourceLines.Add(new MarkdownSourceLineSlice(string.Empty, absoluteLineOffset + index + 1, 1));
                    index++;
                }
                continue;
            }

            if (CountLeadingIndentColumns(line) < continuationIndent) {
                return;
            }

            var stripped = StripLeadingIndentColumns(line, continuationIndent);
            definitionSourceLines.Add(new MarkdownSourceLineSlice(
                stripped,
                absoluteLineOffset + index + 1,
                GetStartColumnAfterStrippingIndent(line, continuationIndent)));
            index++;
        }
    }

    private static (IReadOnlyList<IMarkdownBlock> Blocks, IReadOnlyList<MarkdownSyntaxNode> SyntaxChildren) ParseDefinitionBody(
        List<MarkdownSourceLineSlice> definitionSourceLines,
        MarkdownReaderOptions options,
        MarkdownReaderState state) {
        if (definitionSourceLines == null || definitionSourceLines.Count == 0) {
            return (Array.Empty<IMarkdownBlock>(), Array.Empty<MarkdownSyntaxNode>());
        }

        var (blocks, syntaxChildren) = ParseNestedMarkdownBlocks(definitionSourceLines, options, state);
        if (blocks.Count > 0) {
            return (blocks, syntaxChildren);
        }

        var paragraphs = ParseParagraphBlocksFromSourceLines(definitionSourceLines, options, state);
        var nodes = new List<MarkdownSyntaxNode>();
        AddParagraphSyntaxNodes(nodes, definitionSourceLines, options, state);
        return (paragraphs, nodes);
    }

    private static int GetStartColumnAfterStrippingIndent(string line, int requiredColumns) {
        if (string.IsNullOrEmpty(line) || requiredColumns <= 0) {
            return 1;
        }

        int rawIndex = 0;
        int consumedColumns = 0;
        while (rawIndex < line.Length && consumedColumns < requiredColumns) {
            if (line[rawIndex] == ' ') {
                consumedColumns++;
                rawIndex++;
                continue;
            }

            if (line[rawIndex] == '\t') {
                consumedColumns = ((consumedColumns / 4) + 1) * 4;
                rawIndex++;
                continue;
            }

            break;
        }

        return rawIndex + 1;
    }

    private static string RenderDefinitionLiteral(IReadOnlyList<IMarkdownBlock> blocks, string fallbackLiteral) {
        if (blocks == null || blocks.Count == 0) {
            return fallbackLiteral ?? string.Empty;
        }

        var sb = new StringBuilder();
        for (int i = 0; i < blocks.Count; i++) {
            var rendered = blocks[i]?.RenderMarkdown();
            if (string.IsNullOrEmpty(rendered)) {
                continue;
            }

            if (sb.Length > 0) {
                sb.Append("\n\n");
            }

            sb.Append(rendered);
        }

        return sb.Length > 0 ? sb.ToString() : fallbackLiteral ?? string.Empty;
    }
}

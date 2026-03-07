namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static MarkdownSyntaxNode BuildDocumentSyntaxTree(IReadOnlyList<MarkdownSyntaxNode> children) =>
        new MarkdownSyntaxNode(MarkdownSyntaxKind.Document, GetAggregateSpan(children), children: children);

    private static void CaptureSyntaxNodes(MarkdownDoc doc, int previousBlockCount, int startLine, int endExclusiveLine, List<MarkdownSyntaxNode> nodes) {
        int start = startLine + 1;
        int end = Math.Max(start, endExclusiveLine);
        var span = new MarkdownSourceSpan(start, end);

        for (int blockIndex = previousBlockCount; blockIndex < doc.Blocks.Count; blockIndex++) {
            nodes.Add(BuildSyntaxNode(doc.Blocks[blockIndex], span));
        }
    }

    private static MarkdownSyntaxNode BuildSyntaxNode(IMarkdownBlock block, MarkdownSourceSpan? span = null) {
        switch (block) {
            case HeadingBlock heading:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.Heading, span, heading.Text);
            case ParagraphBlock paragraph:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.Paragraph, span, paragraph.Inlines.RenderMarkdown());
            case HorizontalRuleBlock:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.HorizontalRule, span, "---");
            case CodeBlock code:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.CodeBlock, span, NormalizeSyntaxLiteralLineEndings(code.Content), BuildCodeBlockChildren(code, span));
            case ImageBlock image:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.Image, span, ((IMarkdownBlock)image).RenderMarkdown());
            case TableBlock table:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.Table, span, ((IMarkdownBlock)table).RenderMarkdown(), BuildTableChildren(table, span));
            case QuoteBlock quote:
                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.Quote,
                    span,
                    quote.Children.Count == 0 ? string.Join("\n", quote.Lines) : null,
                    quote.SyntaxChildren ?? BuildChildSyntaxNodes(quote.Children));
            case UnorderedListBlock unordered:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.UnorderedList, span, children: BuildListItemSyntaxNodes(unordered.Items, MarkdownSyntaxKind.UnorderedList));
            case OrderedListBlock ordered:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.OrderedList, span, ordered.Start.ToString(System.Globalization.CultureInfo.InvariantCulture), BuildListItemSyntaxNodes(ordered.Items, MarkdownSyntaxKind.OrderedList));
            case DefinitionListBlock definitionList:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.DefinitionList, span, children: BuildDefinitionItemSyntaxNodes(definitionList));
            case CalloutBlock callout:
                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.Callout,
                    span,
                    string.IsNullOrWhiteSpace(callout.Title) ? callout.Kind : callout.Kind + ":" + callout.Title,
                    callout.SyntaxChildren ?? (callout.Children != null ? BuildChildSyntaxNodes(callout.Children) : Array.Empty<MarkdownSyntaxNode>()));
            case DetailsBlock details:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.Details, span, details.Open ? "open" : null, BuildDetailsChildren(details));
            case SummaryBlock summary:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.Summary, span ?? summary.SyntaxSpan, summary.Inlines.RenderMarkdown());
            case FootnoteDefinitionBlock footnote:
                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.FootnoteDefinition,
                    span,
                    footnote.Label,
                    BuildFootnoteChildren(footnote));
            case FrontMatterBlock frontMatter:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.FrontMatter, span, frontMatter.Render());
            case HtmlCommentBlock comment:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.HtmlComment, span, comment.Comment);
            case HtmlRawBlock rawHtml:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.HtmlRaw, span, rawHtml.Html);
            case TocBlock toc:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.Toc, span, ((IMarkdownBlock)toc).RenderMarkdown());
            case TocPlaceholderBlock:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.TocPlaceholder, span);
            default:
                return new MarkdownSyntaxNode(MarkdownSyntaxKind.Unknown, span, block.RenderMarkdown());
        }
    }

    private static IReadOnlyList<MarkdownSyntaxNode> BuildChildSyntaxNodes(IEnumerable<IMarkdownBlock> children) {
        var nodes = new List<MarkdownSyntaxNode>();
        foreach (var child in children) {
            if (child == null) continue;
            nodes.Add(BuildSyntaxNode(child));
        }
        return nodes;
    }

    private static IReadOnlyList<MarkdownSyntaxNode> BuildListItemSyntaxNodes(IReadOnlyList<ListItem> items, MarkdownSyntaxKind listKind) {
        int index = 0;
        return BuildListItemSyntaxNodes(items, listKind, ref index, 0);
    }

    private static IReadOnlyList<MarkdownSyntaxNode> BuildListItemSyntaxNodes(IReadOnlyList<ListItem> items, MarkdownSyntaxKind listKind, ref int index, int level) {
        var nodes = new List<MarkdownSyntaxNode>();
        while (index < items.Count) {
            var item = items[index];
            if (item.Level < level) break;
            if (item.Level > level) {
                index++;
                continue;
            }

            index++;

            MarkdownSyntaxNode? nestedList = null;
            if (index < items.Count && items[index].Level > level) {
                var nestedLevel = items[index].Level;
                var nestedItems = BuildListItemSyntaxNodes(items, listKind, ref index, nestedLevel);
                nestedList = new MarkdownSyntaxNode(listKind, GetAggregateSpan(nestedItems), children: nestedItems);
            }

            nodes.Add(BuildListItemSyntaxNode(item, nestedList));
        }

        return nodes;
    }

    private static MarkdownSyntaxNode BuildListItemSyntaxNode(ListItem item, MarkdownSyntaxNode? nestedList) {
        var children = new List<MarkdownSyntaxNode>();
        if (item.SyntaxChildren.Count > 0) {
            children.AddRange(item.SyntaxChildren);
        } else {
            if (item.Content.Items.Count > 0 || (item.AdditionalParagraphs.Count == 0 && item.Children.Count == 0)) {
                children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.Paragraph, literal: item.Content.RenderMarkdown()));
            }
            for (int i = 0; i < item.AdditionalParagraphs.Count; i++) {
                children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.Paragraph, literal: item.AdditionalParagraphs[i].RenderMarkdown()));
            }
            for (int i = 0; i < item.Children.Count; i++) {
                children.Add(BuildSyntaxNode(item.Children[i]));
            }
        }
        if (nestedList != null) children.Add(nestedList);

        string? literal = item.IsTask
            ? (item.Checked ? "[x]" : "[ ]")
            : null;

        return new MarkdownSyntaxNode(MarkdownSyntaxKind.ListItem, GetAggregateSpan(children), literal, children);
    }

    private static IReadOnlyList<MarkdownSyntaxNode> BuildDefinitionItemSyntaxNodes(DefinitionListBlock definitionList) {
        if (definitionList.SyntaxItems.Count > 0) return definitionList.SyntaxItems;

        var nodes = new List<MarkdownSyntaxNode>();
        foreach (var (term, definition) in definitionList.Items) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.DefinitionItem,
                literal: term,
                children: new[] {
                    new MarkdownSyntaxNode(MarkdownSyntaxKind.DefinitionTerm, literal: term),
                    new MarkdownSyntaxNode(
                        MarkdownSyntaxKind.DefinitionValue,
                        literal: definition,
                        children: new[] {
                            new MarkdownSyntaxNode(MarkdownSyntaxKind.Paragraph, literal: definition)
                        })
                }));
        }
        return nodes;
    }

    private static IReadOnlyList<MarkdownSyntaxNode> BuildDetailsChildren(DetailsBlock details) {
        if (details.SyntaxChildren != null && details.SyntaxChildren.Count > 0) {
            var nodesWithSummary = new List<MarkdownSyntaxNode>();
            if (details.Summary != null) nodesWithSummary.Add(BuildSyntaxNode(details.Summary));
            for (int i = 0; i < details.SyntaxChildren.Count; i++) nodesWithSummary.Add(details.SyntaxChildren[i]);
            return nodesWithSummary;
        }

        var nodes = new List<MarkdownSyntaxNode>();
        if (details.Summary != null) nodes.Add(BuildSyntaxNode(details.Summary));
        for (int i = 0; i < details.Children.Count; i++) {
            nodes.Add(BuildSyntaxNode(details.Children[i]));
        }
        return nodes;
    }

    private static IReadOnlyList<MarkdownSyntaxNode> BuildFootnoteChildren(FootnoteDefinitionBlock footnote) {
        if (footnote.SyntaxChildren != null && footnote.SyntaxChildren.Count > 0) return footnote.SyntaxChildren;

        if (footnote.Paragraphs.Count == 0) return Array.Empty<MarkdownSyntaxNode>();

        var nodes = new List<MarkdownSyntaxNode>(footnote.Paragraphs.Count);
        for (int i = 0; i < footnote.Paragraphs.Count; i++) {
            nodes.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.Paragraph, literal: footnote.Paragraphs[i].RenderMarkdown()));
        }
        return nodes;
    }

    private static IReadOnlyList<MarkdownSyntaxNode> BuildCodeBlockChildren(CodeBlock code, MarkdownSourceSpan? span) {
        if (!span.HasValue) return Array.Empty<MarkdownSyntaxNode>();

        var nodes = new List<MarkdownSyntaxNode>();
        if (code.IsFenced && !string.IsNullOrEmpty(code.Language)) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.CodeFenceInfo,
                new MarkdownSourceSpan(span.Value.StartLine, span.Value.StartLine),
                code.Language));
        }

        MarkdownSourceSpan? contentSpan;
        if (code.IsFenced) {
            contentSpan = span.Value.EndLine > span.Value.StartLine + 1
                ? new MarkdownSourceSpan(span.Value.StartLine + 1, span.Value.EndLine - 1)
                : null;
        } else {
            contentSpan = span.Value;
        }

        nodes.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.CodeContent, contentSpan, NormalizeSyntaxLiteralLineEndings(code.Content)));
        return nodes;
    }

    private static IReadOnlyList<MarkdownSyntaxNode> BuildTableChildren(TableBlock table, MarkdownSourceSpan? span) {
        if (!span.HasValue) return Array.Empty<MarkdownSyntaxNode>();

        var nodes = new List<MarkdownSyntaxNode>();
        int line = span.Value.StartLine;

        if (table.Headers.Count > 0) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.TableHeader,
                new MarkdownSourceSpan(line, line),
                string.Join(" | ", table.Headers)));
            line += 2; // Skip the alignment row.
        }

        for (int i = 0; i < table.Rows.Count; i++) {
            if (line > span.Value.EndLine) break;

            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.TableRow,
                new MarkdownSourceSpan(line, line),
                string.Join(" | ", table.Rows[i])));
            line++;
        }

        return nodes;
    }

    private static string NormalizeSyntaxLiteralLineEndings(string? value) {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        string normalized = value!;
        return normalized.Replace("\r\n", "\n").Replace('\r', '\n');
    }

    private static MarkdownSourceSpan? GetAggregateSpan(IReadOnlyList<MarkdownSyntaxNode> nodes) {
        if (nodes == null || nodes.Count == 0) return null;

        int? start = null;
        int? end = null;
        for (int i = 0; i < nodes.Count; i++) {
            var span = nodes[i].SourceSpan;
            if (!span.HasValue) continue;

            if (!start.HasValue || span.Value.StartLine < start.Value) start = span.Value.StartLine;
            if (!end.HasValue || span.Value.EndLine > end.Value) end = span.Value.EndLine;
        }

        if (!start.HasValue || !end.HasValue) return null;
        return new MarkdownSourceSpan(start.Value, end.Value);
    }
}

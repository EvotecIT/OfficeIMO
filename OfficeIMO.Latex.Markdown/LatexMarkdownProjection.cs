namespace OfficeIMO.Latex.Markdown;

// One projection owns Markdown, plain text, source locations, and fidelity warnings.
internal sealed class LatexMarkdownProjection {
    internal LatexMarkdownProjection(LatexDocument document, LatexToMarkdownResult result, IReadOnlyList<LatexProjectedBlock> blocks) {
        Document = document;
        Result = result;
        Blocks = blocks;
    }
    internal LatexDocument Document { get; }
    internal LatexToMarkdownResult Result { get; }
    internal IReadOnlyList<LatexProjectedBlock> Blocks { get; }
}

internal sealed class LatexProjectedBlock {
    internal LatexProjectedBlock(LatexSourceSpan span, string kind, MarkdownDoc markdown, IReadOnlyList<LatexMarkdownConversionDiagnostic> diagnostics) {
        Span = span;
        Kind = kind;
        Document = markdown;
        Diagnostics = diagnostics;
    }
    internal LatexSourceSpan Span { get; }
    internal string Kind { get; }
    internal MarkdownDoc Document { get; }
    internal IReadOnlyList<LatexMarkdownConversionDiagnostic> Diagnostics { get; }
    internal string Text => LatexProjectedText.Extract(Document.Blocks);
    internal string Markdown => Document.ToMarkdown().TrimEnd();
}

internal static class LatexProjectedText {
    internal static string Extract(IEnumerable<IMarkdownBlock> blocks) =>
        string.Join("\n\n", blocks.Select(Extract).Where(static text => !string.IsNullOrWhiteSpace(text)));

    private static string Extract(IMarkdownBlock block) {
        switch (block) {
            case HeadingBlock heading: return InlinePlainText.Extract(heading.Inlines);
            case ParagraphBlock paragraph: return InlinePlainText.Extract(paragraph.Inlines);
            case CodeBlock code: return code.Content;
            case SemanticFencedBlock semantic: return semantic.Content;
            case ImageBlock image: return (image.Caption ?? image.PlainAlt ?? string.Empty) + "\n" + image.Path;
            case UnorderedListBlock list: return ExtractItems(list.Items);
            case OrderedListBlock list: return ExtractItems(list.Items);
            case DefinitionListBlock definitions:
                return string.Join("\n", definitions.Entries.Select(entry => InlinePlainText.Extract(entry.Term) + ": " + Extract(entry.DefinitionBlocks)));
            case TableBlock table: {
                var rows = new List<string>();
                string header = string.Join("\t", table.HeaderCells.Select(cell => Extract(cell.ChildBlocks)));
                if (!string.IsNullOrWhiteSpace(header)) rows.Add(header);
                rows.AddRange(table.RowCells.Select(row => string.Join("\t", row.Select(cell => Extract(cell.ChildBlocks)))));
                string? caption = table.Attributes.Attributes.FirstOrDefault(static pair => pair.Key == "caption").Value;
                if (!string.IsNullOrWhiteSpace(caption)) rows.Insert(0, caption!);
                return string.Join("\n", rows);
            }
            case IChildMarkdownBlockContainer container: return Extract(container.ChildBlocks);
            default: return string.Empty;
        }
    }

    private static string ExtractItems(IEnumerable<ListItem> items) =>
        string.Join("\n", items.Select(item => InlinePlainText.Extract(item.Content) + Extract(item.NestedBlocks)));
}

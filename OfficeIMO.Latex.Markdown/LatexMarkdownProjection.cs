namespace OfficeIMO.Latex.Markdown;

// One projection owns Markdown, plain text, source locations, and fidelity warnings.
internal sealed class LatexMarkdownProjection {
    internal LatexMarkdownProjection(LatexDocument document, LatexToMarkdownResult result, IReadOnlyList<LatexProjectedBlock> blocks,
        IReadOnlyList<LatexMarkdownConversionDiagnostic> globalDiagnostics) {
        Document = document;
        Result = result;
        Blocks = blocks;
        GlobalDiagnostics = globalDiagnostics;
    }
    internal LatexDocument Document { get; }
    internal LatexToMarkdownResult Result { get; }
    internal IReadOnlyList<LatexProjectedBlock> Blocks { get; }
    internal IReadOnlyList<LatexMarkdownConversionDiagnostic> GlobalDiagnostics { get; }
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
    internal string Text => Kind == "figure" ? LatexProjectedText.ExtractFigure(Document.Blocks) : LatexProjectedText.Extract(Document.Blocks);
    internal string Markdown => Document.ToMarkdown().TrimEnd();
}

internal static class LatexProjectedText {
    internal static string ExtractFigure(IEnumerable<IMarkdownBlock> blocks) {
        var text = new List<string>();
        var captions = new HashSet<string>(StringComparer.Ordinal);
        foreach (IMarkdownBlock block in blocks) {
            if (block is ImageBlock image) {
                string caption = image.Caption ?? image.PlainAlt ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(caption) && captions.Add(caption)) text.Add(caption);
                text.Add(image.Path);
            } else {
                string value = Extract(block);
                if (!string.IsNullOrWhiteSpace(value)) text.Add(value);
            }
        }
        return string.Join("\n", text);
    }

    internal static string Extract(IEnumerable<IMarkdownBlock> blocks) =>
        string.Join("\n\n", blocks.Select(Extract).Where(static text => !string.IsNullOrWhiteSpace(text)));

    private static string Extract(IMarkdownBlock block) {
        switch (block) {
            case HeadingBlock heading: return ExtractInlines(heading.Inlines);
            case ParagraphBlock paragraph: return ExtractInlines(paragraph.Inlines);
            case CodeBlock code: return code.Content;
            case SemanticFencedBlock semantic: return semantic.Content;
            case ImageBlock image: return (image.Caption ?? image.PlainAlt ?? string.Empty) + "\n" + image.Path;
            case UnorderedListBlock list: return ExtractItems(list.Items);
            case OrderedListBlock list: return ExtractItems(list.Items);
            case DefinitionListBlock definitions:
                return string.Join("\n", definitions.Entries.Select(entry => ExtractInlines(entry.Term) + ": " + Extract(entry.DefinitionBlocks)));
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
        string.Join("\n", items.Select(item => ExtractInlines(item.Content) + Extract(item.NestedBlocks)));

    // Split Reader chunks use this projection as their Markdown carrier too. Keep non-visible
    // label anchors on their own line so ordinary whitespace splitting cannot sever the tag.
    private static string ExtractInlines(InlineSequence? sequence) {
        if (sequence == null) return string.Empty;
        if (sequence.Nodes.Count == 1) {
            if (sequence.Nodes[0] is MarkdownTextRun text) return text.Text;
            if (sequence.Nodes[0] is CodeSpanInline code) return code.Text;
        }
        var output = new StringBuilder();
        foreach (IMarkdownInline inline in sequence.Nodes) {
            if (inline is HtmlRawInline html) output.Append('\n').Append(html.Html).Append('\n');
            else if (inline is IInlineContainerMarkdownInline container) output.Append(ExtractInlines(container.NestedInlines));
            else if (inline is InlineSequence nested) output.Append(ExtractInlines(nested));
            else ((IPlainTextMarkdownInline)inline).AppendPlainText(output);
        }
        return output.ToString();
    }
}

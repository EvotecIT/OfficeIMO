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
    internal IReadOnlyList<LatexTextSegment> TextSegments(System.Threading.CancellationToken cancellationToken) =>
        LatexProjectedText.Segments(Document.Blocks, Kind == "figure", cancellationToken);
    internal string Markdown => Document.ToMarkdown().TrimEnd();
}

// Preserve the provenance of non-visible anchors and opaque payloads when Reader splits text.
internal sealed class LatexTextSegment {
    internal LatexTextSegment(string text, string? language = null, bool anchor = false) {
        Text = text; Language = language; IsAnchor = anchor;
    }
    internal string Text { get; }
    internal string? Language { get; }
    internal bool IsAnchor { get; }
}

internal static class LatexProjectedText {
    internal static string ExtractFigure(IEnumerable<IMarkdownBlock> blocks) => Text(Segments(blocks, true));
    internal static string Extract(IEnumerable<IMarkdownBlock> blocks) => Text(Segments(blocks));
    internal static string Text(IEnumerable<LatexTextSegment> segments) => string.Concat(segments.Select(static segment => segment.Text));

    internal static IReadOnlyList<LatexTextSegment> Segments(IEnumerable<IMarkdownBlock> blocks, bool figure = false,
        System.Threading.CancellationToken cancellationToken = default) {
        var parts = new List<IReadOnlyList<LatexTextSegment>>();
        var captions = new HashSet<string>(StringComparer.Ordinal);
        foreach (IMarkdownBlock block in blocks) {
            cancellationToken.ThrowIfCancellationRequested();
            if (figure && block is ImageBlock image) {
                string caption = image.Caption ?? image.PlainAlt ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(caption) && captions.Add(caption)) parts.Add(Plain(caption));
                parts.Add(Plain(image.Path));
            } else parts.Add(Block(block, cancellationToken));
        }
        return Join(parts, figure ? "\n" : "\n\n");
    }

    private static IReadOnlyList<LatexTextSegment> Block(IMarkdownBlock block, System.Threading.CancellationToken token) {
        token.ThrowIfCancellationRequested();
        switch (block) {
            case HeadingBlock heading: return Inlines(heading.Inlines, token);
            case ParagraphBlock paragraph: return Inlines(paragraph.Inlines, token);
            case CodeBlock code: return new[] { new LatexTextSegment(code.Content, code.Language) };
            case SemanticFencedBlock semantic: return new[] { new LatexTextSegment(semantic.Content, semantic.Language) };
            case ImageBlock image: return Plain((image.Caption ?? image.PlainAlt ?? string.Empty) + "\n" + image.Path);
            case UnorderedListBlock list: return Items(list.Items, token);
            case OrderedListBlock list: return Items(list.Items, token);
            case DefinitionListBlock definitions:
                return Join(definitions.Entries.Select(entry => Definition(entry, token)), "\n");
            case TableBlock table: {
                var rows = new List<IReadOnlyList<LatexTextSegment>>();
                string? caption = table.Attributes.Attributes.FirstOrDefault(static pair => pair.Key == "caption").Value;
                if (!string.IsNullOrWhiteSpace(caption)) rows.Add(Plain(caption!));
                IReadOnlyList<LatexTextSegment> header = Join(table.HeaderCells.Select(cell => Segments(cell.ChildBlocks, cancellationToken: token)), "\t", false);
                if (!string.IsNullOrWhiteSpace(Text(header))) rows.Add(header);
                rows.AddRange(table.RowCells.Select(row => Join(row.Select(cell => Segments(cell.ChildBlocks, cancellationToken: token)), "\t", false)));
                return Join(rows, "\n", false);
            }
            case IChildMarkdownBlockContainer container: return Segments(container.ChildBlocks, cancellationToken: token);
            default: return Array.Empty<LatexTextSegment>();
        }
    }

    private static IReadOnlyList<LatexTextSegment> Items(IEnumerable<ListItem> items, System.Threading.CancellationToken token) =>
        Join(items.Select(item => (IReadOnlyList<LatexTextSegment>)Inlines(item.Content, token)
            .Concat(Segments(item.NestedBlocks, cancellationToken: token)).ToArray()), "\n");

    internal static bool HasVisibleText(InlineSequence sequence, System.Threading.CancellationToken token) =>
        HasVisibleText(Inlines(sequence, token), token);

    private static bool HasVisibleText(IEnumerable<LatexTextSegment> segments, System.Threading.CancellationToken token) {
        foreach (LatexTextSegment segment in segments) {
            token.ThrowIfCancellationRequested();
            if (segment.IsAnchor) continue;
            for (int index = 0; index < segment.Text.Length; index++) {
                if ((index & 1023) == 0) token.ThrowIfCancellationRequested();
                if (!char.IsWhiteSpace(segment.Text[index])) return true;
            }
        }
        return false;
    }

    private static IReadOnlyList<LatexTextSegment> Definition(DefinitionListEntry entry, System.Threading.CancellationToken token) {
        IReadOnlyList<LatexTextSegment> term = Inlines(entry.Term, token);
        IReadOnlyList<LatexTextSegment> body = Segments(entry.DefinitionBlocks, cancellationToken: token);
        bool visibleTerm = HasVisibleText(term, token);
        if (!visibleTerm && !term.Any(static segment => segment.IsAnchor)) term = Array.Empty<LatexTextSegment>();
        return term.Concat(visibleTerm && HasVisibleText(body, token) ? Plain(": ") : Array.Empty<LatexTextSegment>()).Concat(body).ToArray();
    }

    private static IReadOnlyList<LatexTextSegment> Inlines(InlineSequence? sequence, System.Threading.CancellationToken token) {
        var parts = new List<LatexTextSegment>();
        if (sequence == null) return parts;
        foreach (IMarkdownInline inline in sequence.Nodes) {
            token.ThrowIfCancellationRequested();
            if (inline is HtmlRawInline html) {
                parts.Add(new LatexTextSegment("\n"));
                parts.Add(new LatexTextSegment(html.Html, anchor: true));
                parts.Add(new LatexTextSegment("\n"));
            } else if (inline is CodeSpanInline code) parts.Add(new LatexTextSegment(code.Text, "text"));
            else if (inline is IInlineContainerMarkdownInline container && container.NestedInlines != null)
                parts.AddRange(Inlines(container.NestedInlines, token));
            else if (inline is InlineSequence nested) parts.AddRange(Inlines(nested, token));
            else {
                var text = new StringBuilder();
                ((IPlainTextMarkdownInline)inline).AppendPlainText(text);
                parts.Add(new LatexTextSegment(text.ToString()));
            }
        }
        return parts;
    }

    private static IReadOnlyList<LatexTextSegment> Plain(string text) => new[] { new LatexTextSegment(text) };
    private static IReadOnlyList<LatexTextSegment> Join(IEnumerable<IReadOnlyList<LatexTextSegment>> parts, string separator, bool omitWhitespace = true) {
        var result = new List<LatexTextSegment>();
        bool first = true;
        foreach (IReadOnlyList<LatexTextSegment> part in parts) {
            if (omitWhitespace && part.All(static segment => string.IsNullOrWhiteSpace(segment.Text))) continue;
            if (!first) result.Add(new LatexTextSegment(separator));
            result.AddRange(part);
            first = false;
        }
        return result;
    }
}

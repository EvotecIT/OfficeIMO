namespace OfficeIMO.Markdown;

/// <summary>
/// Simple blockquote block consisting of raw text lines.
/// </summary>
public sealed class QuoteBlock : MarkdownBlock, IMarkdownBlock, IChildMarkdownBlockContainer, ISyntaxChildrenMarkdownBlock, IOwnedSyntaxChildrenMarkdownBlock, ISyntaxMarkdownBlock {
    private readonly List<MarkdownSourceSpan> _markerSourceSpans = new List<MarkdownSourceSpan>();

    /// <summary>Raw text lines for a simple quote (used when <see cref="ChildBlocks"/> is empty).</summary>
    public System.Collections.Generic.List<string> Lines { get; } = new System.Collections.Generic.List<string>();
    /// <summary>Nested blocks rendered inside the quote.</summary>
    public System.Collections.Generic.List<IMarkdownBlock> ChildBlocks { get; } = new System.Collections.Generic.List<IMarkdownBlock>();
    /// <summary>Source spans for quote marker tokens (<c>&gt;</c>) captured from parsed markdown lines.</summary>
    public IReadOnlyList<MarkdownSourceSpan> MarkerSourceSpans => _markerSourceSpans;
    /// <summary>Nested syntax nodes captured during parsing, when available.</summary>
    internal IReadOnlyList<MarkdownSyntaxNode>? SyntaxChildren { get; set; }
    /// <summary>Create an empty quote block.</summary>
    public QuoteBlock() { }
    /// <summary>Create a quote block with initial lines.</summary>
    public QuoteBlock(System.Collections.Generic.IEnumerable<string> lines) { Lines.AddRange(lines); }

    internal void ClearSyntaxCache() {
        SyntaxChildren = null;
    }

    internal void ReplaceMarkerSourceSpans(IEnumerable<MarkdownSourceSpan>? spans) {
        _markerSourceSpans.Clear();
        if (spans == null) {
            return;
        }

        _markerSourceSpans.AddRange(spans);
    }

    string IMarkdownBlock.RenderMarkdown() {
        if (ChildBlocks.Count > 0) {
            var sb = new StringBuilder();
            for (int i = 0; i < ChildBlocks.Count; i++) {
                var rendered = MarkdownBlockRenderDispatcher.RenderMarkdown(ChildBlocks[i]);
                // Prefix every line with "> "
                using var reader = new System.IO.StringReader(rendered);
                string? line; bool first = true;
                while ((line = reader.ReadLine()) != null) {
                    if (!first) sb.AppendLine();
                    sb.Append("> ").Append(EscapeQuoteLineForMarkdown(line));
                    first = false;
                }
                if (i < ChildBlocks.Count - 1) sb.AppendLine().AppendLine("> "); // blank quote line to separate blocks
            }
            return sb.ToString();
        }
        var sb2 = new StringBuilder();
        foreach (var l in Lines) sb2.AppendLine("> " + EscapeQuoteLineForMarkdown(l));
        return sb2.ToString().TrimEnd();
    }

    private static string EscapeQuoteLineForMarkdown(string? line) {
        if (string.IsNullOrEmpty(line)) {
            return line ?? string.Empty;
        }

        var value = line ?? string.Empty;
        return MarkdownReader.TryGetSetextHeadingUnderlineLevel(value, out _)
            ? "\\" + value
            : value;
    }

    string IMarkdownBlock.RenderHtml() {
        if (ChildBlocks.Count > 0) {
            var sb = new StringBuilder();
            sb.Append("<blockquote>");
            foreach (var b in ChildBlocks) {
                var rendered = MarkdownBlockRenderDispatcher.RenderHtml(b);
                if (RequiresRawHtmlBlockBoundary(rendered, b)) {
                    sb.AppendLine();
                    sb.Append(rendered);
                    if (!EndsWithLineBreak(rendered)) {
                        sb.AppendLine();
                    }
                    continue;
                }

                sb.Append(rendered);
            }
            sb.Append("</blockquote>");
            return sb.ToString();
        }
        if (Lines.Count == 0) {
            return "<blockquote></blockquote>";
        }

        var encoded = HtmlTextEncoder.Encode(string.Join("\n", Lines), HtmlRenderContext.Options).Replace("\n", "<br/>");
        return $"<blockquote><p>{encoded}</p></blockquote>";
    }

    private static bool RequiresRawHtmlBlockBoundary(string rendered, IMarkdownBlock block) {
        if (string.IsNullOrEmpty(rendered)) {
            return false;
        }

        return block is HtmlRawBlock or HtmlCommentBlock;
    }

    private static bool EndsWithLineBreak(string value) {
        return value.EndsWith("\n", StringComparison.Ordinal) ||
               value.EndsWith("\r", StringComparison.Ordinal);
    }

    IReadOnlyList<IMarkdownBlock> IChildMarkdownBlockContainer.ChildBlocks => ChildBlocks;
    IReadOnlyList<MarkdownSyntaxNode>? ISyntaxChildrenMarkdownBlock.ProvidedSyntaxChildren => SyntaxChildren;

    IReadOnlyList<MarkdownSyntaxNode> IOwnedSyntaxChildrenMarkdownBlock.BuildOwnedSyntaxChildren() {
        return MarkdownBlockSyntaxBuilder.BuildCanonicalChildSyntaxNodes(SyntaxChildren, ChildBlocks);
    }

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) {
        var children = BuildQuoteMarkerSyntaxNodes();
        children.AddRange(((IOwnedSyntaxChildrenMarkdownBlock)this).BuildOwnedSyntaxChildren());

        return new MarkdownSyntaxNode(
            MarkdownSyntaxKind.Quote,
            span,
            ChildBlocks.Count == 0 ? string.Join("\n", Lines) : null,
            children,
            this);
    }

    private List<MarkdownSyntaxNode> BuildQuoteMarkerSyntaxNodes() {
        var children = new List<MarkdownSyntaxNode>(_markerSourceSpans.Count);
        for (var i = 0; i < _markerSourceSpans.Count; i++) {
            children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.QuoteMarker, _markerSourceSpans[i], ">"));
        }

        return children;
    }
}

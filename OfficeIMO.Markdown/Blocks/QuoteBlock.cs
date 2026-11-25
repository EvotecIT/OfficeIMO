namespace OfficeIMO.Markdown;

/// <summary>
/// Simple blockquote block consisting of raw text lines.
/// </summary>
public sealed class QuoteBlock : MarkdownBlock, IMarkdownBlock, IChildMarkdownBlockContainer, ISyntaxChildrenMarkdownBlock, IOwnedSyntaxChildrenMarkdownBlock, ISyntaxMarkdownBlock, IHtmlAppendableMarkdownBlock {
    private readonly List<MarkdownSourceSpan> _markerSourceSpans = new List<MarkdownSourceSpan>();

    /// <summary>Raw text lines for a simple quote (used when <see cref="ChildBlocks"/> is empty).</summary>
    public System.Collections.Generic.List<string> Lines { get; } = new System.Collections.Generic.List<string>();
    /// <summary>Nested blocks rendered inside the quote.</summary>
    public System.Collections.Generic.List<IMarkdownBlock> ChildBlocks { get; } = new System.Collections.Generic.List<IMarkdownBlock>();
    /// <summary>Source spans for quote marker tokens (<c>&gt;</c>) captured from parsed markdown lines.</summary>
    public IReadOnlyList<MarkdownSourceSpan> MarkerSourceSpans {
        get {
            EnsureSourceInformation();
            return _markerSourceSpans;
        }
    }
    /// <summary>Nested syntax nodes captured during parsing, when available.</summary>
    internal IReadOnlyList<MarkdownSyntaxNode>? SyntaxChildren { get => ReadSourceInformation(ref field); set => WriteSourceInformation(ref field, value); }
    /// <summary>Create an empty quote block.</summary>
    public QuoteBlock() { }
    /// <summary>Create a quote block with initial lines.</summary>
    public QuoteBlock(System.Collections.Generic.IEnumerable<string> lines) { Lines.AddRange(lines); }

    internal void ClearSyntaxCache() {
        SyntaxChildren = null;
    }

    internal void ReplaceMarkerSourceSpans(IEnumerable<MarkdownSourceSpan>? spans) {
        EnsureSourceInformation();
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
        if (ChildBlocks.Count == 0 && Lines.Count == 0) {
            return "<blockquote></blockquote>";
        }

        var sb = new StringBuilder();
        AppendHtml(sb);
        return sb.ToString();
    }

    void IHtmlAppendableMarkdownBlock.AppendHtml(StringBuilder builder) => AppendHtml(builder);

    private void AppendHtml(StringBuilder sb) {
        if (ChildBlocks.Count > 0) {
            sb.Append("<blockquote>");
            foreach (var b in ChildBlocks) {
                // Raw HTML children need their own lines; everything else is written in place.
                if (b is HtmlRawBlock or HtmlCommentBlock) {
                    var rendered = MarkdownBlockRenderDispatcher.RenderHtml(b);
                    if (RequiresRawHtmlBlockBoundary(rendered, b)) {
                        sb.AppendLine();
                        sb.Append(rendered);
                        if (!EndsWithLineBreak(rendered)) {
                            sb.AppendLine();
                        }
                    }

                    continue;
                }

                MarkdownBlockRenderDispatcher.AppendHtml(sb, b);
            }
            sb.Append("</blockquote>");
            return;
        }
        if (Lines.Count == 0) {
            sb.Append("<blockquote></blockquote>");
            return;
        }

        var encoded = HtmlTextEncoder.Encode(string.Join("\n", Lines), HtmlRenderContext.Options).Replace("\n", "<br/>");
        sb.Append("<blockquote><p>").Append(encoded).Append("</p></blockquote>");
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

    internal override void AdoptCapturedSourceInformation(MarkdownObject captured, MarkdownCapturedObjectMap map) {
        var source = (QuoteBlock)captured;
        _markerSourceSpans.Clear();
        _markerSourceSpans.AddRange(source._markerSourceSpans);
        SyntaxChildren = map.Retarget(source.SyntaxChildren);
    }
}

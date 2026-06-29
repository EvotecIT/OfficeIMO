namespace OfficeIMO.Markdown;

/// <summary>
/// Paragraph block containing a sequence of inline nodes.
/// </summary>
public sealed class ParagraphBlock : MarkdownBlock, IMarkdownBlock, IParagraphMarkdownBlock, IInlineSyntaxMarkdownBlock, ISyntaxMarkdownBlock {
    /// <summary>Inline content within this paragraph.</summary>
    public InlineSequence Inlines { get; }
    /// <summary>Creates a paragraph block.</summary>
    public ParagraphBlock(InlineSequence inlines) { Inlines = inlines; }
    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() => Inlines.RenderMarkdown() + MarkdownAttributeBlockRenderer.RenderTrailing(Attributes);
    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() => $"<p{MarkdownHtmlAttributes.Render(Attributes, null)}>{Inlines.RenderHtml()}</p>";
    InlineSequence IParagraphMarkdownBlock.ParagraphInlines => Inlines;
    string ITightListItemHtmlMarkdownBlock.RenderTightListItemHtml() => Inlines.RenderHtml();
    InlineSequence IInlineSyntaxMarkdownBlock.SyntaxInlines => Inlines;
    MarkdownSyntaxKind IInlineSyntaxMarkdownBlock.SyntaxKind => MarkdownSyntaxKind.Paragraph;
    MarkdownSourceSpan? IInlineSyntaxMarkdownBlock.ProvidedSyntaxSpan => null;
    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        MarkdownBlockSyntaxBuilder.BuildInlineBlock(this, span);
}

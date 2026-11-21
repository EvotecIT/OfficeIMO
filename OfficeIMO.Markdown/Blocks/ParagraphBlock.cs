namespace OfficeIMO.Markdown;

/// <summary>
/// MarkdownParagraph block containing a sequence of inline nodes.
/// </summary>
public sealed class ParagraphBlock : MarkdownBlock, IMarkdownBlock, IParagraphMarkdownBlock, IInlineSyntaxMarkdownBlock, ISyntaxMarkdownBlock, IHtmlAppendableMarkdownBlock {
    /// <summary>Inline content within this paragraph.</summary>
    public InlineSequence Inlines { get; }
    internal string GenericAttributeConsumedWhitespace { get; set; } = string.Empty;
    internal bool GenericAttributeSuppressSeparator { get; set; }
    /// <summary>Creates a paragraph block.</summary>
    public ParagraphBlock(InlineSequence inlines) { Inlines = inlines; }
    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() {
        if (Attributes.IsEmpty) {
            return Inlines.RenderMarkdown();
        }

        var separator = GenericAttributeSuppressSeparator
            ? string.Empty
            : string.IsNullOrEmpty(GenericAttributeConsumedWhitespace)
                ? " "
                : GenericAttributeConsumedWhitespace;
        return Inlines.RenderMarkdown() + separator + MarkdownAttributeBlockRenderer.RenderInlineTrailing(Attributes);
    }
    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() {
        var sb = new StringBuilder();
        AppendHtml(sb);
        return sb.ToString();
    }

    void IHtmlAppendableMarkdownBlock.AppendHtml(StringBuilder builder) => AppendHtml(builder);

    private void AppendHtml(StringBuilder sb) {
        sb.Append("<p").Append(MarkdownHtmlAttributes.Render(Attributes, null)).Append('>');
        Inlines.AppendHtml(sb);
        sb.Append(RenderGenericAttributeConsumedWhitespace()).Append("</p>");
    }
    InlineSequence IParagraphMarkdownBlock.ParagraphInlines => Inlines;
    string ITightListItemHtmlMarkdownBlock.RenderTightListItemHtml() => Inlines.RenderHtml();
    InlineSequence IInlineSyntaxMarkdownBlock.SyntaxInlines => Inlines;
    MarkdownSyntaxKind IInlineSyntaxMarkdownBlock.SyntaxKind => MarkdownSyntaxKind.Paragraph;
    MarkdownSourceSpan? IInlineSyntaxMarkdownBlock.ProvidedSyntaxSpan => null;
    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        MarkdownBlockSyntaxBuilder.BuildInlineBlock(this, span);

    private string RenderGenericAttributeConsumedWhitespace() {
        if (string.IsNullOrEmpty(GenericAttributeConsumedWhitespace) || Attributes.IsEmpty) {
            return string.Empty;
        }

        return HtmlTextEncoder.Encode(GenericAttributeConsumedWhitespace, HtmlRenderContext.Options);
    }
}

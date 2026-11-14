namespace OfficeIMO.Markdown;

/// <summary>
/// Writes an inline's default HTML into the caller's buffer, so nested inlines render
/// without a string per node. <see cref="IRenderableMarkdownInline.RenderHtml"/> stays the public form.
/// </summary>
internal interface IHtmlAppendableMarkdownInline {
    void AppendHtml(StringBuilder builder);
}

/// <summary>
/// Writes a block's default HTML into the caller's buffer, so a document renders into one builder.
/// <see cref="IMarkdownBlock.RenderHtml"/> stays the public form.
/// </summary>
internal interface IHtmlAppendableMarkdownBlock {
    void AppendHtml(StringBuilder builder);
}

/// <summary>
/// Writes a context-aware block's HTML into the caller's buffer.
/// <see cref="IContextualHtmlMarkdownBlock.RenderHtml"/> stays the public form.
/// </summary>
internal interface IHtmlAppendableContextualMarkdownBlock {
    void AppendHtml(StringBuilder builder, MarkdownBodyRenderContext context);
}

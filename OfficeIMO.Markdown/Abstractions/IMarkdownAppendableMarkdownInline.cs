namespace OfficeIMO.Markdown;

/// <summary>
/// Writes an inline's default Markdown into the caller's buffer, so nested inlines write without a
/// string per node. <see cref="IRenderableMarkdownInline.RenderMarkdown"/> stays the public form.
/// </summary>
internal interface IMarkdownAppendableMarkdownInline {
    void AppendMarkdown(StringBuilder builder);
}
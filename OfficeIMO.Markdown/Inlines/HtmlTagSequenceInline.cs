namespace OfficeIMO.Markdown;

/// <summary>
/// Inline node rendered via a whitelisted HTML tag while preserving nested inline content as AST.
/// </summary>
public sealed class HtmlTagSequenceInline : IMarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline, IInlineContainerMarkdownInline {
    /// <summary>Normalized lowercase tag name.</summary>
    public string TagName { get; }

    /// <summary>Nested inline content.</summary>
    public InlineSequence Inlines { get; }

    /// <summary>Creates a new HTML-tag inline wrapper.</summary>
    public HtmlTagSequenceInline(string tagName, InlineSequence? inlines = null) {
        if (string.IsNullOrWhiteSpace(tagName)) {
            throw new ArgumentException("Tag name is required.", nameof(tagName));
        }

        TagName = tagName.Trim().ToLowerInvariant();
        Inlines = inlines ?? new InlineSequence();
    }

    internal string RenderMarkdown() => "<" + TagName + ">" + Inlines.RenderMarkdown() + "</" + TagName + ">";
    internal string RenderHtml() => "<" + TagName + ">" + Inlines.RenderHtml() + "</" + TagName + ">";

    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => InlinePlainText.AppendPlainText(sb, Inlines);
    InlineSequence? IInlineContainerMarkdownInline.NestedInlines => Inlines;
}

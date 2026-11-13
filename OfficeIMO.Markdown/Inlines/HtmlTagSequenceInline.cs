namespace OfficeIMO.Markdown;

/// <summary>
/// Inline node rendered via a whitelisted HTML tag while preserving nested inline content as AST.
/// </summary>
public sealed class HtmlTagSequenceInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline, IInlineContainerMarkdownInline, IMarkdownInlineAuxiliarySyntaxMetadataOwner, IHtmlAppendableMarkdownInline {
    MarkdownInlineAuxiliarySyntaxMetadata? IMarkdownInlineAuxiliarySyntaxMetadataOwner.AuxiliarySyntaxMetadata { get; set; }
    /// <summary>Normalized lowercase tag name.</summary>
    public string TagName { get; }

    /// <summary>Nested inline content.</summary>
    public InlineSequence Inlines { get; }

    /// <summary>Creates a new HTML-tag inline wrapper.</summary>
    public HtmlTagSequenceInline(string tagName, InlineSequence? inlines = null) {
        ArgumentException.ThrowIfNullOrWhiteSpace(tagName);

        TagName = tagName.Trim().ToLowerInvariant();
        Inlines = inlines ?? new InlineSequence();
    }

    internal string RenderMarkdown() => "<" + TagName + ">" + Inlines.RenderMarkdown() + "</" + TagName + ">";
    internal string RenderHtml() {
        var sb = new StringBuilder();
        AppendHtml(sb);
        return sb.ToString();
    }
    internal void AppendHtml(StringBuilder sb) {
        sb.Append('<').Append(TagName).Append('>');
        Inlines.AppendHtml(sb);
        sb.Append("</").Append(TagName).Append('>');
    }

    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IHtmlAppendableMarkdownInline.AppendHtml(StringBuilder builder) => AppendHtml(builder);
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => InlinePlainText.AppendPlainText(sb, Inlines);
    InlineSequence? IInlineContainerMarkdownInline.NestedInlines => Inlines;
}

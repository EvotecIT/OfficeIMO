namespace OfficeIMO.Markdown;

/// <summary>
/// Italic/emphasis that contains nested inline nodes.
/// Used by the reader so nested markup can be represented without changing the fluent builder API.
/// </summary>
public sealed class ItalicSequenceInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline, IInlineContainerMarkdownInline, IMarkdownInlineAuxiliarySyntaxMetadataOwner, IHtmlAppendableMarkdownInline {
    MarkdownInlineAuxiliarySyntaxMetadata? IMarkdownInlineAuxiliarySyntaxMetadataOwner.AuxiliarySyntaxMetadata { get; set; }
    /// <summary>Inline content.</summary>
    public InlineSequence Inlines { get; }

    /// <summary>Creates an italic inline with nested inline content.</summary>
    public ItalicSequenceInline(InlineSequence inlines) {
        Inlines = inlines ?? new InlineSequence();
    }

    internal string RenderMarkdown() => "*" + Inlines.RenderMarkdown() + "*";
    internal string RenderHtml() {
        var sb = new StringBuilder();
        AppendHtml(sb);
        return sb.ToString();
    }
    internal void AppendHtml(StringBuilder sb) {
        sb.Append("<em>");
        Inlines.AppendHtml(sb);
        sb.Append("</em>");
    }
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IHtmlAppendableMarkdownInline.AppendHtml(StringBuilder builder) => AppendHtml(builder);
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => InlinePlainText.AppendPlainText(sb, Inlines);
    InlineSequence? IInlineContainerMarkdownInline.NestedInlines => Inlines;
}


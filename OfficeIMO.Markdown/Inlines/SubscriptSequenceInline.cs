namespace OfficeIMO.Markdown;

/// <summary>
/// Subscript inline content that can contain nested inline nodes.
/// Used by Markdig-style emphasis extras so nested markup can be represented without flattening formatting.
/// </summary>
public sealed class SubscriptSequenceInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline, IInlineContainerMarkdownInline, IMarkdownInlineAuxiliarySyntaxMetadataOwner, IHtmlAppendableMarkdownInline {
    MarkdownInlineAuxiliarySyntaxMetadata? IMarkdownInlineAuxiliarySyntaxMetadataOwner.AuxiliarySyntaxMetadata { get; set; }
    /// <summary>Inline content.</summary>
    public InlineSequence Inlines { get; }

    /// <summary>Creates a subscript inline with nested inline content.</summary>
    public SubscriptSequenceInline(InlineSequence inlines) {
        Inlines = inlines ?? new InlineSequence();
    }

    internal string RenderMarkdown() => "~" + Inlines.RenderMarkdownWithTextEscaper(MarkdownEscaper.EscapeSubscriptText) + "~";
    internal string RenderHtml() {
        var sb = new StringBuilder();
        AppendHtml(sb);
        return sb.ToString();
    }
    internal void AppendHtml(StringBuilder sb) {
        sb.Append("<sub>");
        Inlines.AppendHtml(sb);
        sb.Append("</sub>");
    }
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IHtmlAppendableMarkdownInline.AppendHtml(StringBuilder builder) => AppendHtml(builder);
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => InlinePlainText.AppendPlainText(sb, Inlines);
    InlineSequence? IInlineContainerMarkdownInline.NestedInlines => Inlines;
}

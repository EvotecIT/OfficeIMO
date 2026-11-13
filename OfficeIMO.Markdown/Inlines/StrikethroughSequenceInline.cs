namespace OfficeIMO.Markdown;

/// <summary>
/// Strikethrough that contains nested inline nodes.
/// Used by the reader so nested markup can be represented without changing the fluent builder API.
/// </summary>
public sealed class StrikethroughSequenceInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline, IInlineContainerMarkdownInline, IMarkdownInlineAuxiliarySyntaxMetadataOwner, IHtmlAppendableMarkdownInline {
    MarkdownInlineAuxiliarySyntaxMetadata? IMarkdownInlineAuxiliarySyntaxMetadataOwner.AuxiliarySyntaxMetadata { get; set; }
    /// <summary>Inline content.</summary>
    public InlineSequence Inlines { get; }

    /// <summary>Creates a strikethrough inline with nested inline content.</summary>
    public StrikethroughSequenceInline(InlineSequence inlines) {
        Inlines = inlines ?? new InlineSequence();
    }

    internal string RenderMarkdown() => "~~" + Inlines.RenderMarkdown() + "~~";
    internal string RenderHtml() {
        var sb = new StringBuilder();
        AppendHtml(sb);
        return sb.ToString();
    }
    internal void AppendHtml(StringBuilder sb) {
        sb.Append("<del>");
        Inlines.AppendHtml(sb);
        sb.Append("</del>");
    }
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IHtmlAppendableMarkdownInline.AppendHtml(StringBuilder builder) => AppendHtml(builder);
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => InlinePlainText.AppendPlainText(sb, Inlines);
    InlineSequence? IInlineContainerMarkdownInline.NestedInlines => Inlines;
}


namespace OfficeIMO.Markdown;

/// <summary>
/// Superscript inline content that can contain nested inline nodes.
/// Used by Markdig-style emphasis extras so nested markup can be represented without flattening formatting.
/// </summary>
public sealed class SuperscriptSequenceInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline, IInlineContainerMarkdownInline {
    /// <summary>Inline content.</summary>
    public InlineSequence Inlines { get; }

    /// <summary>Creates a superscript inline with nested inline content.</summary>
    public SuperscriptSequenceInline(InlineSequence inlines) {
        Inlines = inlines ?? new InlineSequence();
    }

    internal string RenderMarkdown() => "^" + Inlines.RenderMarkdownWithTextEscaper(MarkdownEscaper.EscapeSuperscriptText) + "^";
    internal string RenderHtml() => "<sup>" + Inlines.RenderHtml() + "</sup>";
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => InlinePlainText.AppendPlainText(sb, Inlines);
    InlineSequence? IInlineContainerMarkdownInline.NestedInlines => Inlines;
}

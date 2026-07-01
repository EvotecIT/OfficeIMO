namespace OfficeIMO.Markdown;

/// <summary>
/// Abbreviation inline rendered as an HTML <c>&lt;abbr&gt;</c> element.
/// </summary>
public sealed class AbbreviationInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline {
    /// <summary>Visible abbreviation text.</summary>
    public string Text { get; }

    /// <summary>Expanded title shown by HTML renderers.</summary>
    public string Title { get; }

    /// <summary>Source span for the visible abbreviation text when parsed from markdown.</summary>
    public MarkdownSourceSpan? TextSourceSpan { get; internal set; }

    /// <summary>Source span for the abbreviation definition title used by this inline.</summary>
    public MarkdownSourceSpan? TitleSourceSpan { get; internal set; }

    /// <summary>Creates an abbreviation inline.</summary>
    public AbbreviationInline(string text, string title) {
        Text = text ?? string.Empty;
        Title = title ?? string.Empty;
    }

    internal void SetMarkdownSyntaxMetadataSpans(
        MarkdownSourceSpan? textSourceSpan,
        MarkdownSourceSpan? titleSourceSpan) {
        TextSourceSpan = textSourceSpan;
        TitleSourceSpan = titleSourceSpan;
    }

    internal string RenderMarkdown() => MarkdownEscaper.EscapeText(Text);

    internal string RenderHtml() =>
        "<abbr title=\"" + HtmlTextEncoder.Encode(Title, HtmlRenderContext.Options) + "\">" +
        HtmlTextEncoder.Encode(Text, HtmlRenderContext.Options) +
        "</abbr>";

    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();

    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();

    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(Text);
}

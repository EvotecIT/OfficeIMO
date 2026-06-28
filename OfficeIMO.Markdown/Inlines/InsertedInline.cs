namespace OfficeIMO.Markdown;

/// <summary>
/// Inserted inline text rendered as <c>++text++</c> in Markdown and <c>&lt;ins&gt;</c> in HTML.
/// </summary>
public sealed class InsertedInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline {
    /// <summary>Text content.</summary>
    public string Text { get; }

    /// <summary>Creates a new inserted inline.</summary>
    public InsertedInline(string text) {
        Text = text ?? string.Empty;
    }

    internal string RenderMarkdown() => "++" + MarkdownEscaper.EscapeEmphasis(Text) + "++";
    internal string RenderHtml() => "<ins>" + System.Net.WebUtility.HtmlEncode(Text) + "</ins>";
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(Text);
}

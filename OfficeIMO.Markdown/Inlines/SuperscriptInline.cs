namespace OfficeIMO.Markdown;

/// <summary>
/// Superscript inline text rendered as <c>^text^</c> in Markdown and <c>&lt;sup&gt;</c> in HTML.
/// </summary>
public sealed class SuperscriptInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline {
    /// <summary>Text content.</summary>
    public string Text { get; }

    /// <summary>Creates a new superscript inline.</summary>
    public SuperscriptInline(string text) {
        Text = text ?? string.Empty;
    }

    internal string RenderMarkdown() => "^" + MarkdownEscaper.EscapeSuperscriptText(Text) + "^";
    internal string RenderHtml() => "<sup>" + System.Net.WebUtility.HtmlEncode(Text) + "</sup>";
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(Text);
}

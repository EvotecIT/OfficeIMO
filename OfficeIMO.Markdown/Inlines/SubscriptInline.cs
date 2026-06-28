namespace OfficeIMO.Markdown;

/// <summary>
/// Subscript inline text rendered as <c>~text~</c> in Markdown and <c>&lt;sub&gt;</c> in HTML.
/// </summary>
public sealed class SubscriptInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline {
    /// <summary>Text content.</summary>
    public string Text { get; }

    /// <summary>Creates a new subscript inline.</summary>
    public SubscriptInline(string text) {
        Text = text ?? string.Empty;
    }

    internal string RenderMarkdown() => "~" + MarkdownEscaper.EscapeSubscriptText(Text) + "~";
    internal string RenderHtml() => "<sub>" + System.Net.WebUtility.HtmlEncode(Text) + "</sub>";
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(Text);
}

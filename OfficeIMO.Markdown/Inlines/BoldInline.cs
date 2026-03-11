namespace OfficeIMO.Markdown;

/// <summary>
/// Bold/strong emphasis inline.
/// </summary>
public sealed class BoldInline : IMarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline, IStrongMarkdownInline {
    /// <summary>Text content.</summary>
    public string Text { get; }
    /// <summary>Creates a bold inline with the given text.</summary>
    public BoldInline(string text) { Text = text ?? string.Empty; }
    internal string RenderMarkdown() => "**" + MarkdownEscaper.EscapeEmphasis(Text) + "**";
    internal string RenderHtml() => "<strong>" + System.Net.WebUtility.HtmlEncode(Text) + "</strong>";
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(Text);
}

namespace OfficeIMO.Markdown;

/// <summary>
/// Abbreviation inline rendered as an HTML <c>&lt;abbr&gt;</c> element.
/// </summary>
public sealed class AbbreviationInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline {
    /// <summary>Visible abbreviation text.</summary>
    public string Text { get; }

    /// <summary>Expanded title shown by HTML renderers.</summary>
    public string Title { get; }

    /// <summary>Creates an abbreviation inline.</summary>
    public AbbreviationInline(string text, string title) {
        Text = text ?? string.Empty;
        Title = title ?? string.Empty;
    }

    internal string RenderMarkdown() => MarkdownEscaper.EscapeText(Text);

    internal string RenderHtml() =>
        "<abbr title=\"" + System.Net.WebUtility.HtmlEncode(Title) + "\">" +
        System.Net.WebUtility.HtmlEncode(Text) +
        "</abbr>";

    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();

    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();

    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(Text);
}

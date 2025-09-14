namespace OfficeIMO.Markdown;

/// <summary>
/// Footnote reference inline, e.g., [^1].
/// </summary>
public sealed class FootnoteRefInline {
    public string Label { get; }
    public FootnoteRefInline(string label) { Label = label ?? string.Empty; }
    internal string RenderMarkdown() => $"[^{Label}]";
    internal string RenderHtml() => $"<sup id=\"fnref:{System.Net.WebUtility.HtmlEncode(Label)}\"><a href=\"#fn:{System.Net.WebUtility.HtmlEncode(Label)}\">{System.Net.WebUtility.HtmlEncode(Label)}</a></sup>";
}


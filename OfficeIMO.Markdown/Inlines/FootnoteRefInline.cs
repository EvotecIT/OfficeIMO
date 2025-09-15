namespace OfficeIMO.Markdown;

/// <summary>
/// Footnote reference inline, e.g., [^1].
/// </summary>
public sealed class FootnoteRefInline {
    /// <summary>Reference label pointing to a footnote definition.</summary>
    public string Label { get; }
    /// <summary>Create a footnote reference inline.</summary>
    public FootnoteRefInline(string label) { Label = label ?? string.Empty; }
    internal string RenderMarkdown() => $"[^{Label}]";
    internal string RenderHtml() => $"<sup id=\"fnref:{System.Net.WebUtility.HtmlEncode(Label)}\"><a href=\"#fn:{System.Net.WebUtility.HtmlEncode(Label)}\">{System.Net.WebUtility.HtmlEncode(Label)}</a></sup>";
}

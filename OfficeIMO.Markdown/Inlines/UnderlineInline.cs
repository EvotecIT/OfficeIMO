namespace OfficeIMO.Markdown;

/// <summary>
/// Underline inline. Not native to CommonMark; we render as <u> in Markdown (HTML passthrough) and HTML.
/// </summary>
public sealed class UnderlineInline {
    /// <summary>Text to underline.</summary>
    public string Text { get; }
    public UnderlineInline(string text) { Text = text ?? string.Empty; }
    internal string RenderMarkdown() => $"<u>{System.Net.WebUtility.HtmlEncode(Text)}</u>";
    internal string RenderHtml() => $"<u>{System.Net.WebUtility.HtmlEncode(Text)}</u>";
}


namespace OfficeIMO.Markdown;

/// <summary>
/// Underline inline. Not native to CommonMark; we render as &lt;u&gt; in Markdown (HTML passthrough) and HTML.
/// </summary>
public sealed class UnderlineInline {
    /// <summary>Text to underline.</summary>
    public string Text { get; }
    /// <summary>
    /// Creates a new underline inline with the provided text.
    /// </summary>
    /// <param name="text">Text to render underlined.</param>
    public UnderlineInline(string text) { Text = text ?? string.Empty; }
    internal string RenderMarkdown() => $"<u>{System.Net.WebUtility.HtmlEncode(Text)}</u>";
    internal string RenderHtml() => $"<u>{System.Net.WebUtility.HtmlEncode(Text)}</u>";
}

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Information about a rendered fenced code block matched by the renderer extension pipeline.
/// </summary>
public sealed class MarkdownFencedCodeBlockMatch {
    /// <summary>
    /// Creates a new fenced code block match payload.
    /// </summary>
    public MarkdownFencedCodeBlockMatch(string language, string htmlEncodedContent, string rawContent, string originalHtml) {
        Language = language ?? string.Empty;
        HtmlEncodedContent = htmlEncodedContent ?? string.Empty;
        RawContent = rawContent ?? string.Empty;
        OriginalHtml = originalHtml ?? string.Empty;
    }

    /// <summary>
    /// Language / info string captured from the rendered <c>language-*</c> class.
    /// </summary>
    public string Language { get; }

    /// <summary>
    /// HTML-encoded code contents as emitted by the markdown HTML renderer.
    /// </summary>
    public string HtmlEncodedContent { get; }

    /// <summary>
    /// HTML-decoded raw code contents.
    /// </summary>
    public string RawContent { get; }

    /// <summary>
    /// Original HTML fragment for the matched <c>&lt;pre&gt;&lt;code&gt;</c> block.
    /// </summary>
    public string OriginalHtml { get; }
}

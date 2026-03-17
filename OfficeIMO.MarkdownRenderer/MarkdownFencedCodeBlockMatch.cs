using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Information about a rendered fenced code block matched by the renderer extension pipeline.
/// </summary>
public sealed class MarkdownFencedCodeBlockMatch {
    /// <summary>
    /// Creates a new fenced code block match payload.
    /// </summary>
    public MarkdownFencedCodeBlockMatch(string infoString, string htmlEncodedContent, string rawContent, string originalHtml) {
        FenceInfo = MarkdownCodeFenceInfo.Parse(infoString);
        InfoString = FenceInfo.InfoString;
        Language = FenceInfo.Language;
        HtmlEncodedContent = htmlEncodedContent ?? string.Empty;
        RawContent = rawContent ?? string.Empty;
        OriginalHtml = originalHtml ?? string.Empty;
    }

    /// <summary>
    /// Parsed primary fence language token.
    /// </summary>
    public string Language { get; }

    /// <summary>
    /// Full fenced-code info string preserved from the AST/source block.
    /// </summary>
    public string InfoString { get; }

    /// <summary>
    /// Structured fenced-code info metadata.
    /// </summary>
    public MarkdownCodeFenceInfo FenceInfo { get; }

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

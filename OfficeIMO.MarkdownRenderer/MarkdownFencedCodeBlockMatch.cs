using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Information about a rendered fenced code block matched by the renderer extension pipeline.
/// </summary>
public sealed class MarkdownFencedCodeBlockMatch {
    /// <summary>
    /// Creates a new fenced code block match payload.
    /// </summary>
    public MarkdownFencedCodeBlockMatch(string infoString, string htmlEncodedContent, string rawContent, string originalHtml)
        : this(infoString, htmlEncodedContent, rawContent, originalHtml, null, null) {
    }

    /// <summary>Creates a match with the original fence and content source locations.</summary>
    public MarkdownFencedCodeBlockMatch(string infoString, string htmlEncodedContent, string rawContent, string originalHtml,
        MarkdownSourceSpan? sourceSpan, MarkdownSourceSpan? contentSourceSpan) {
        FenceInfo = MarkdownCodeFenceInfo.Parse(infoString);
        InfoString = FenceInfo.InfoString;
        Language = FenceInfo.Language;
        HtmlEncodedContent = htmlEncodedContent ?? string.Empty;
        RawContent = rawContent ?? string.Empty;
        OriginalHtml = originalHtml ?? string.Empty;
        SourceSpan = sourceSpan;
        ContentSourceSpan = contentSourceSpan;
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

    /// <summary>Original source location of the fence, when available.</summary>
    public MarkdownSourceSpan? SourceSpan { get; }

    /// <summary>Original source location of the fence contents, when available.</summary>
    public MarkdownSourceSpan? ContentSourceSpan { get; }
}

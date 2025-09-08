namespace OfficeIMO.Markdown;

/// <summary>
/// Convenience block for heading level 3.
/// </summary>
public sealed class H3 : IMarkdownBlock {
    private readonly HeadingBlock _h;
    /// <summary>Creates an H3 heading.</summary>
    public H3(string text) { _h = new HeadingBlock(3, text); }
    /// <inheritdoc />
    public string RenderMarkdown() => _h.RenderMarkdown();
    /// <inheritdoc />
    public string RenderHtml() => _h.RenderHtml();
}

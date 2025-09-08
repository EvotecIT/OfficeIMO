namespace OfficeIMO.Markdown;

/// <summary>
/// Convenience block for heading level 1.
/// </summary>
public sealed class H1 : IMarkdownBlock {
    private readonly HeadingBlock _h;
    /// <summary>Creates an H1 heading.</summary>
    public H1(string text) { _h = new HeadingBlock(1, text); }
    /// <inheritdoc />
    public string RenderMarkdown() => _h.RenderMarkdown();
    /// <inheritdoc />
    public string RenderHtml() => _h.RenderHtml();
}

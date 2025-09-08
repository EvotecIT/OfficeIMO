namespace OfficeIMO.Markdown;

/// <summary>
/// Convenience block for heading level 2.
/// </summary>
public sealed class H2 : IMarkdownBlock {
    private readonly HeadingBlock _h;
    /// <summary>Creates an H2 heading.</summary>
    public H2(string text) { _h = new HeadingBlock(2, text); }
    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() => ((IMarkdownBlock)_h).RenderMarkdown();
    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() => ((IMarkdownBlock)_h).RenderHtml();
}

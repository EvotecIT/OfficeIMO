namespace OfficeIMO.Markdown;

/// <summary>
/// Convenience block for heading level 2.
/// </summary>
public sealed class MarkdownH2 : IMarkdownBlock {
    private readonly HeadingBlock _h;
    /// <summary>Creates an MarkdownH2 heading.</summary>
    public MarkdownH2(string text) { _h = new HeadingBlock(2, text); }
    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() => ((IMarkdownBlock)_h).RenderMarkdown();
    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() => ((IMarkdownBlock)_h).RenderHtml();
}

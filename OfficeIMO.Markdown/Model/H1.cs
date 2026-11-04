namespace OfficeIMO.Markdown;

/// <summary>
/// Convenience block for heading level 1.
/// </summary>
public sealed class MarkdownH1 : IMarkdownBlock {
    private readonly HeadingBlock _h;
    /// <summary>Creates an MarkdownH1 heading.</summary>
    public MarkdownH1(string text) { _h = new HeadingBlock(1, text); }
    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() => ((IMarkdownBlock)_h).RenderMarkdown();
    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() => ((IMarkdownBlock)_h).RenderHtml();
}

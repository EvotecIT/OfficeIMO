namespace OfficeIMO.Markdown;

/// <summary>
/// Placeholder that is replaced with a generated Table of Contents at render time.
/// </summary>
internal sealed class TocPlaceholderBlock : IMarkdownBlock {
    public TocOptions Options { get; }
    public TocPlaceholderBlock(TocOptions options) { Options = options; }
    string IMarkdownBlock.RenderMarkdown() => string.Empty; // Replaced during render
    string IMarkdownBlock.RenderHtml() => string.Empty; // Replaced during render
}


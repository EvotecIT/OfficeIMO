namespace OfficeIMO.Markdown;

/// <summary>
/// Context available while rendering a markdown inline node to HTML.
/// </summary>
public sealed class MarkdownInlineHtmlRenderContext {
    private readonly MarkdownBodyRenderContext? _bodyContext;

    internal MarkdownInlineHtmlRenderContext(HtmlOptions options, MarkdownBodyRenderContext? bodyContext) {
        Options = options ?? throw new ArgumentNullException(nameof(options));
        _bodyContext = bodyContext;
    }

    /// <summary>
    /// Active HTML rendering options.
    /// </summary>
    public HtmlOptions Options { get; }

    /// <summary>
    /// Top-level blocks being rendered for the current body, or an empty list when no body context is active.
    /// </summary>
    public IReadOnlyList<IMarkdownBlock> Blocks => _bodyContext?.Blocks ?? Array.Empty<IMarkdownBlock>();

    /// <summary>
    /// Returns the zero-based index of a top-level block in <see cref="Blocks"/>, or <c>-1</c> when unavailable.
    /// </summary>
    public int GetBlockIndex(IMarkdownBlock block) => _bodyContext?.GetBlockIndex(block) ?? -1;

    /// <summary>
    /// Returns the resolved anchor id for a heading block within the current rendered body.
    /// </summary>
    public string GetHeadingAnchor(IMarkdownBlock heading) => _bodyContext?.GetHeadingAnchor(heading) ?? string.Empty;

    /// <summary>
    /// Returns the anchor id of the nearest preceding heading according to the supplied TOC options.
    /// </summary>
    public string? GetPrecedingHeadingAnchor(int blockIndex, TocOptions options) =>
        _bodyContext?.GetPrecedingHeadingAnchor(blockIndex, options);

    /// <summary>
    /// Builds TOC-style heading entries relative to a specific top-level block index using the supplied TOC options.
    /// </summary>
    public IReadOnlyList<TocBlock.Entry> BuildTocEntries(int blockIndex, TocOptions options, string? titleAnchor = null) =>
        _bodyContext?.BuildTocEntries(blockIndex, options, titleAnchor) ?? Array.Empty<TocBlock.Entry>();
}


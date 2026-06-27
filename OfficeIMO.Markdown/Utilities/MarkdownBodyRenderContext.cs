namespace OfficeIMO.Markdown;

/// <summary>
/// Context available while rendering a markdown document body to HTML.
/// </summary>
public sealed class MarkdownBodyRenderContext {
    internal MarkdownBodyRenderContext(
        MarkdownDoc document,
        IReadOnlyList<IMarkdownBlock> blocks,
        HtmlOptions options,
        MarkdownHeadingCatalog headingCatalog) {
        Document = document ?? throw new ArgumentNullException(nameof(document));
        Blocks = blocks;
        Options = options;
        HeadingCatalog = headingCatalog;
    }

    /// <summary>
    /// Document being rendered.
    /// </summary>
    public MarkdownDoc Document { get; }

    /// <summary>
    /// Top-level blocks being rendered for the current body.
    /// </summary>
    public IReadOnlyList<IMarkdownBlock> Blocks { get; }

    /// <summary>
    /// Active HTML rendering options.
    /// </summary>
    public HtmlOptions Options { get; }

    internal MarkdownHeadingCatalog HeadingCatalog { get; }

    /// <summary>
    /// Returns the zero-based index of a top-level block in <see cref="Blocks"/>, or <c>-1</c> when the block is not present.
    /// </summary>
    public int GetBlockIndex(IMarkdownBlock block) {
        if (block == null) {
            return -1;
        }

        for (int i = 0; i < Blocks.Count; i++) {
            if (ReferenceEquals(Blocks[i], block)) {
                return i;
            }
        }

        return -1;
    }

    /// <summary>
    /// Returns the resolved anchor id for a heading block within the current rendered body.
    /// </summary>
    public string GetHeadingAnchor(IMarkdownBlock heading) =>
        heading is IHeadingMarkdownBlock headingBlock
            ? HeadingCatalog.GetHeadingAnchor(headingBlock)
            : string.Empty;

    /// <summary>
    /// Returns the anchor id of the nearest preceding heading according to the supplied TOC options,
    /// or <c>null</c> when no heading title should be associated with the specified block index.
    /// </summary>
    public string? GetPrecedingHeadingAnchor(int blockIndex, TocOptions options) =>
        HeadingCatalog.GetPrecedingHeadingAnchor(Blocks, blockIndex, options ?? new TocOptions());

    /// <summary>
    /// Builds TOC-style heading entries relative to a specific top-level block index using the supplied TOC options.
    /// </summary>
    public IReadOnlyList<TocBlock.Entry> BuildTocEntries(int blockIndex, TocOptions options, string? titleAnchor = null) =>
        HeadingCatalog.BuildTocEntries(Blocks, blockIndex, options ?? new TocOptions(), titleAnchor);

    /// <summary>
    /// Renders a block through the active HTML render dispatcher, including syntax-kind and type-targeted overrides.
    /// Contextual custom container blocks should use this for child blocks instead of calling <see cref="IMarkdownBlock.RenderHtml"/> directly.
    /// </summary>
    public string RenderBlock(IMarkdownBlock block) =>
        MarkdownBlockRenderDispatcher.RenderHtml(block, this);

    /// <summary>
    /// Finds the final syntax-tree node associated with a parsed model object, or <c>null</c> for builder-only documents.
    /// </summary>
    public MarkdownSyntaxNode? FindSyntaxNode(object associatedObject) =>
        Document.ParseResult?.FindFinalNodeForAssociatedObject(associatedObject);

    /// <summary>
    /// Creates a normalized source slice for the final syntax node associated with a parsed model object.
    /// </summary>
    public bool TryCreateSourceSlice(object associatedObject, out MarkdownSourceSlice slice) {
        var parseResult = Document.ParseResult;
        if (parseResult == null) {
            slice = default;
            return false;
        }

        return parseResult.TryCreateSourceSlice(associatedObject, out slice);
    }

    /// <summary>
    /// Creates a normalized source slice for the supplied final syntax node.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownSyntaxNode syntaxNode, out MarkdownSourceSlice slice) {
        var parseResult = Document.ParseResult;
        if (parseResult == null) {
            slice = default;
            return false;
        }

        return parseResult.TryCreateSourceSlice(syntaxNode, out slice);
    }

    /// <summary>
    /// Creates an original-input source slice for the final syntax node associated with a parsed model object.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(object associatedObject, out MarkdownSourceSlice slice) {
        var parseResult = Document.ParseResult;
        if (parseResult == null) {
            slice = default;
            return false;
        }

        return parseResult.TryCreateOriginalSourceSlice(associatedObject, out slice);
    }

    /// <summary>
    /// Creates an original-input source slice for the supplied final syntax node.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownSyntaxNode syntaxNode, out MarkdownSourceSlice slice) {
        var parseResult = Document.ParseResult;
        if (parseResult == null) {
            slice = default;
            return false;
        }

        return parseResult.TryCreateOriginalSourceSlice(syntaxNode, out slice);
    }
}

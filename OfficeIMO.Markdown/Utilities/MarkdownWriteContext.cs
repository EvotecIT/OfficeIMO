namespace OfficeIMO.Markdown;

/// <summary>
/// Context available while rendering a markdown document body back to Markdown text.
/// </summary>
public sealed class MarkdownWriteContext {
    internal MarkdownWriteContext(
        MarkdownDoc document,
        IReadOnlyList<IMarkdownBlock> blocks,
        MarkdownWriteOptions options,
        MarkdownHeadingCatalog headingCatalog) {
        Document = document ?? throw new ArgumentNullException(nameof(document));
        Blocks = blocks;
        Options = options;
        HeadingCatalog = headingCatalog;
    }

    /// <summary>
    /// Document being serialized.
    /// </summary>
    public MarkdownDoc Document { get; }

    /// <summary>
    /// Top-level blocks being rendered for the current markdown document.
    /// </summary>
    public IReadOnlyList<IMarkdownBlock> Blocks { get; }

    /// <summary>
    /// Active markdown writer options.
    /// </summary>
    public MarkdownWriteOptions Options { get; }

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
    /// Returns the resolved anchor id for a heading block within the current rendered document.
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
    /// Renders a block through the active Markdown writer dispatcher, including syntax-kind and type-targeted overrides.
    /// Custom container renderers should use this for child blocks instead of calling <see cref="IMarkdownBlock.RenderMarkdown"/> directly.
    /// </summary>
    public string RenderBlock(IMarkdownBlock block) =>
        MarkdownBlockRenderDispatcher.RenderMarkdown(block, this);

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
    /// Creates a normalized source slice for a token or field source span captured during parsing.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownSourceSpan sourceSpan, out MarkdownSourceSlice slice) {
        var parseResult = Document.ParseResult;
        if (parseResult == null) {
            slice = default;
            return false;
        }

        return parseResult.TryCreateSourceSlice(sourceSpan, out slice);
    }

    /// <summary>
    /// Creates an original-input source slice for the final syntax node associated with a parsed model object.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(object associatedObject, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(associatedObject, out slice, out _);
    }

    /// <summary>
    /// Creates an original-input source slice for the final syntax node associated with a parsed model object.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        object associatedObject,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        var parseResult = Document.ParseResult;
        if (parseResult == null) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved;
            return false;
        }

        return parseResult.TryCreateOriginalSourceSlice(associatedObject, out slice, out failureReason);
    }

    /// <summary>
    /// Creates an original-input source slice for the supplied final syntax node.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownSyntaxNode syntaxNode, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(syntaxNode, out slice, out _);
    }

    /// <summary>
    /// Creates an original-input source slice for the supplied final syntax node.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownSyntaxNode syntaxNode,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        var parseResult = Document.ParseResult;
        if (parseResult == null) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved;
            return false;
        }

        return parseResult.TryCreateOriginalSourceSlice(syntaxNode, out slice, out failureReason);
    }

    /// <summary>
    /// Creates an original-input source slice for a token or field source span captured during parsing.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownSourceSpan sourceSpan, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(sourceSpan, out slice, out _);
    }

    /// <summary>
    /// Creates an original-input source slice for a token or field source span captured during parsing.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownSourceSpan sourceSpan,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        var parseResult = Document.ParseResult;
        if (parseResult == null) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved;
            return false;
        }

        return parseResult.TryCreateOriginalSourceSlice(sourceSpan, out slice, out failureReason);
    }
}

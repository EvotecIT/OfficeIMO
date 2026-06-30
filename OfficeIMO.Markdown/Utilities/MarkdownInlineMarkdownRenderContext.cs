namespace OfficeIMO.Markdown;

/// <summary>
/// Context available while rendering a markdown inline node back to Markdown text.
/// </summary>
public sealed class MarkdownInlineMarkdownRenderContext {
    private readonly MarkdownWriteContext? _writeContext;

    internal MarkdownInlineMarkdownRenderContext(MarkdownWriteOptions options, MarkdownWriteContext? writeContext) {
        Options = options ?? throw new ArgumentNullException(nameof(options));
        _writeContext = writeContext;
    }

    /// <summary>
    /// Active markdown writer options.
    /// </summary>
    public MarkdownWriteOptions Options { get; }

    /// <summary>
    /// Top-level blocks being rendered for the current document, or an empty list when no document context is active.
    /// </summary>
    public IReadOnlyList<IMarkdownBlock> Blocks => _writeContext?.Blocks ?? Array.Empty<IMarkdownBlock>();

    /// <summary>
    /// Returns the zero-based index of a top-level block in <see cref="Blocks"/>, or <c>-1</c> when unavailable.
    /// </summary>
    public int GetBlockIndex(IMarkdownBlock block) => _writeContext?.GetBlockIndex(block) ?? -1;

    /// <summary>
    /// Returns the resolved anchor id for a heading block within the current rendered document.
    /// </summary>
    public string GetHeadingAnchor(IMarkdownBlock heading) => _writeContext?.GetHeadingAnchor(heading) ?? string.Empty;

    /// <summary>
    /// Returns the anchor id of the nearest preceding heading according to the supplied TOC options.
    /// </summary>
    public string? GetPrecedingHeadingAnchor(int blockIndex, TocOptions options) =>
        _writeContext?.GetPrecedingHeadingAnchor(blockIndex, options);

    /// <summary>
    /// Builds TOC-style heading entries relative to a specific top-level block index using the supplied TOC options.
    /// </summary>
    public IReadOnlyList<TocBlock.Entry> BuildTocEntries(int blockIndex, TocOptions options, string? titleAnchor = null) =>
        _writeContext?.BuildTocEntries(blockIndex, options, titleAnchor) ?? Array.Empty<TocBlock.Entry>();

    /// <summary>
    /// Finds the final syntax-tree node associated with a parsed model object, or <c>null</c> when no document context is active.
    /// </summary>
    public MarkdownSyntaxNode? FindSyntaxNode(object associatedObject) =>
        _writeContext?.FindSyntaxNode(associatedObject);

    /// <summary>
    /// Creates a normalized source slice for the final syntax node associated with a parsed model object.
    /// </summary>
    public bool TryCreateSourceSlice(object associatedObject, out MarkdownSourceSlice slice) {
        if (_writeContext == null) {
            slice = default;
            return false;
        }

        return _writeContext.TryCreateSourceSlice(associatedObject, out slice);
    }

    /// <summary>
    /// Creates a normalized source slice for the supplied final syntax node.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownSyntaxNode syntaxNode, out MarkdownSourceSlice slice) {
        if (_writeContext == null) {
            slice = default;
            return false;
        }

        return _writeContext.TryCreateSourceSlice(syntaxNode, out slice);
    }

    /// <summary>
    /// Creates a normalized source slice for a token or field source span captured during parsing.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownSourceSpan sourceSpan, out MarkdownSourceSlice slice) {
        if (_writeContext == null) {
            slice = default;
            return false;
        }

        return _writeContext.TryCreateSourceSlice(sourceSpan, out slice);
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
        if (_writeContext == null) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved;
            return false;
        }

        return _writeContext.TryCreateOriginalSourceSlice(associatedObject, out slice, out failureReason);
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
        if (_writeContext == null) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved;
            return false;
        }

        return _writeContext.TryCreateOriginalSourceSlice(syntaxNode, out slice, out failureReason);
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
        if (_writeContext == null) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved;
            return false;
        }

        return _writeContext.TryCreateOriginalSourceSlice(sourceSpan, out slice, out failureReason);
    }
}

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

    /// <summary>
    /// Encodes text content with the active HTML escaping policy.
    /// </summary>
    public string EncodeText(string? text) => HtmlTextEncoder.Encode(text, Options);

    /// <summary>
    /// Encodes a quoted HTML attribute value with the active HTML escaping policy.
    /// </summary>
    public string EncodeAttributeValue(string? value) => HtmlTextEncoder.Encode(value, Options);

    /// <summary>
    /// Encodes a URL-bearing HTML attribute value with the active URL and HTML escaping policy.
    /// </summary>
    public string EncodeUrlAttribute(string? url) => HtmlAttributeUrlEncoder.Encode(url, Options);

    /// <summary>
    /// Finds the final syntax-tree node associated with a parsed model object, or <c>null</c> when no body context is active.
    /// </summary>
    public MarkdownSyntaxNode? FindSyntaxNode(object associatedObject) =>
        _bodyContext?.FindSyntaxNode(associatedObject);

    /// <summary>
    /// Creates a normalized source slice for the final syntax node associated with a parsed model object.
    /// </summary>
    public bool TryCreateSourceSlice(object associatedObject, out MarkdownSourceSlice slice) {
        if (_bodyContext == null) {
            slice = default;
            return false;
        }

        return _bodyContext.TryCreateSourceSlice(associatedObject, out slice);
    }

    /// <summary>
    /// Creates a normalized source slice for the supplied final syntax node.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownSyntaxNode syntaxNode, out MarkdownSourceSlice slice) {
        if (_bodyContext == null) {
            slice = default;
            return false;
        }

        return _bodyContext.TryCreateSourceSlice(syntaxNode, out slice);
    }

    /// <summary>
    /// Creates a normalized source slice for a token or field source span captured during parsing.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownSourceSpan sourceSpan, out MarkdownSourceSlice slice) {
        if (_bodyContext == null) {
            slice = default;
            return false;
        }

        return _bodyContext.TryCreateSourceSlice(sourceSpan, out slice);
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
        if (_bodyContext == null) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved;
            return false;
        }

        return _bodyContext.TryCreateOriginalSourceSlice(associatedObject, out slice, out failureReason);
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
        if (_bodyContext == null) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved;
            return false;
        }

        return _bodyContext.TryCreateOriginalSourceSlice(syntaxNode, out slice, out failureReason);
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
        if (_bodyContext == null) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved;
            return false;
        }

        return _bodyContext.TryCreateOriginalSourceSlice(sourceSpan, out slice, out failureReason);
    }
}

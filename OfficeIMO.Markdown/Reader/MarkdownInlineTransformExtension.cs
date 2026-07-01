namespace OfficeIMO.Markdown;

/// <summary>
/// Context passed to post-parse inline transform extensions.
/// </summary>
public sealed class MarkdownInlineTransformContext {
    private readonly MarkdownReaderState? _state;

    internal MarkdownInlineTransformContext(
        string sourceText,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        bool isNestedSequence) {
        SourceText = sourceText ?? string.Empty;
        Options = options ?? throw new ArgumentNullException(nameof(options));
        _state = state;
        IsNestedSequence = isNestedSequence;
    }

    /// <summary>
    /// Inline source text for the root inline parse that produced the current sequence.
    /// For nested inline containers this remains the root inline source slice.
    /// </summary>
    public string SourceText { get; }

    /// <summary>Reader options active for the current parse.</summary>
    public MarkdownReaderOptions Options { get; }

    /// <summary>
    /// Returns <see langword="true"/> when the transform is visiting a nested inline container sequence.
    /// </summary>
    public bool IsNestedSequence { get; }

    /// <summary>
    /// Returns the source span associated with an inline node, when the node came from parsed Markdown source.
    /// </summary>
    public MarkdownSourceSpan? GetSourceSpan(IMarkdownInline inline) =>
        MarkdownInlineSourceSpans.Get(inline);

    /// <summary>
    /// Creates a normalized source slice for an inline node, when the node came from parsed Markdown source.
    /// </summary>
    public bool TryCreateSourceSlice(IMarkdownInline inline, out MarkdownSourceSlice slice) {
        var span = GetSourceSpan(inline);
        if (!span.HasValue) {
            slice = default;
            return false;
        }

        return TryCreateSourceSlice(span.Value, out slice);
    }

    /// <summary>
    /// Creates a normalized source slice for a token or field source span captured during inline parsing.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownSourceSpan sourceSpan, out MarkdownSourceSlice slice) {
        var sourceMap = _state?.SourceTextMap;
        if (sourceMap == null) {
            slice = default;
            return false;
        }

        return MarkdownSourceSlice.TryCreate(sourceMap.Text, sourceSpan, MarkdownSourceTextKind.Normalized, out slice);
    }
}

/// <summary>
/// Delegate used by post-parse inline transform extensions.
/// Implementations may mutate <paramref name="sequence"/> directly or return a replacement sequence.
/// Returning <see langword="null"/> leaves the current sequence unchanged.
/// </summary>
public delegate InlineSequence? MarkdownInlineTransform(InlineSequence sequence, MarkdownInlineTransformContext context);

/// <summary>
/// Named inline AST transform registration used by <see cref="MarkdownReader"/>.
/// </summary>
public sealed class MarkdownInlineTransformExtension {
    /// <summary>
    /// Creates an inline transform extension registration.
    /// </summary>
    public MarkdownInlineTransformExtension(
        string name,
        MarkdownInlineTransform transform,
        Func<MarkdownReaderOptions, bool>? isEnabled = null) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Extension name cannot be null or whitespace.", nameof(name));
        }

        Name = name.Trim();
        Transform = transform ?? throw new ArgumentNullException(nameof(transform));
        IsEnabled = isEnabled;
    }

    /// <summary>Stable extension name used for diagnostics and inspection.</summary>
    public string Name { get; }

    /// <summary>Transform delegate contributed by this extension.</summary>
    public MarkdownInlineTransform Transform { get; }

    /// <summary>Optional predicate that decides whether the extension should apply for a specific options instance.</summary>
    public Func<MarkdownReaderOptions, bool>? IsEnabled { get; }

    /// <summary>Returns <see langword="true"/> when the extension should participate in inline transformation.</summary>
    public bool AppliesTo(MarkdownReaderOptions options) => IsEnabled?.Invoke(options) ?? true;
}

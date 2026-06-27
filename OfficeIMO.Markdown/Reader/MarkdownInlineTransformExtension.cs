namespace OfficeIMO.Markdown;

/// <summary>
/// Context passed to post-parse inline transform extensions.
/// </summary>
public sealed class MarkdownInlineTransformContext {
    internal MarkdownInlineTransformContext(
        string sourceText,
        MarkdownReaderOptions options,
        bool isNestedSequence) {
        SourceText = sourceText ?? string.Empty;
        Options = options ?? throw new ArgumentNullException(nameof(options));
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

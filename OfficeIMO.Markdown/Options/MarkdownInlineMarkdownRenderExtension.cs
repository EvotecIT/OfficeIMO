namespace OfficeIMO.Markdown;

/// <summary>
/// Delegate used to override Markdown emitted for a specific inline type during Markdown serialization.
/// Returning <see langword="null"/> falls back to the inline's built-in Markdown rendering.
/// </summary>
public delegate string? MarkdownInlineMarkdownRenderer(IMarkdownInline inline, MarkdownWriteOptions options);

/// <summary>
/// Named Markdown render extension that can override emitted Markdown for a specific inline type.
/// </summary>
public sealed class MarkdownInlineMarkdownRenderExtension {
    /// <summary>Creates a Markdown inline render extension registration.</summary>
    public MarkdownInlineMarkdownRenderExtension(
        string name,
        Type inlineType,
        MarkdownInlineMarkdownRenderer renderMarkdown) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Extension name cannot be null or whitespace.", nameof(name));
        }

        Name = name.Trim();
        InlineType = inlineType ?? throw new ArgumentNullException(nameof(inlineType));
        RenderMarkdown = renderMarkdown ?? throw new ArgumentNullException(nameof(renderMarkdown));
    }

    /// <summary>Stable extension name used for inspection or de-duplication.</summary>
    public string Name { get; }

    /// <summary>The inline type this extension handles.</summary>
    public Type InlineType { get; }

    /// <summary>Markdown rendering callback.</summary>
    public MarkdownInlineMarkdownRenderer RenderMarkdown { get; }

    /// <summary>Returns <see langword="true"/> when the extension can render the provided inline node.</summary>
    public bool Matches(IMarkdownInline inline) => inline != null && InlineType.IsInstanceOfType(inline);
}

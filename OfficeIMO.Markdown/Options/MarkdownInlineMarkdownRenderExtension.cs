namespace OfficeIMO.Markdown;

/// <summary>
/// Delegate used to override Markdown emitted for a specific inline type during Markdown serialization.
/// Returning <see langword="null"/> falls back to the inline's built-in Markdown rendering.
/// </summary>
public delegate string? MarkdownInlineMarkdownRenderer(IMarkdownInline inline, MarkdownWriteOptions options);

/// <summary>
/// Delegate used to override Markdown emitted for a specific inline type during Markdown serialization with access
/// to the current markdown write context.
/// Returning <see langword="null"/> falls back to the inline's built-in Markdown rendering.
/// </summary>
public delegate string? MarkdownInlineMarkdownContextualRenderer(IMarkdownInline inline, MarkdownInlineMarkdownRenderContext context);

/// <summary>
/// Named Markdown render extension that can override emitted Markdown for a specific inline type.
/// </summary>
public sealed class MarkdownInlineMarkdownRenderExtension {
    private MarkdownInlineMarkdownRenderExtension(
        string name,
        Type inlineType,
        MarkdownInlineMarkdownRenderer renderMarkdown,
        MarkdownInlineMarkdownContextualRenderer renderMarkdownWithContext) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Extension name cannot be null or whitespace.", nameof(name));
        }

        Name = name.Trim();
        InlineType = inlineType ?? throw new ArgumentNullException(nameof(inlineType));
        RenderMarkdown = renderMarkdown ?? throw new ArgumentNullException(nameof(renderMarkdown));
        RenderMarkdownWithContext = renderMarkdownWithContext ?? throw new ArgumentNullException(nameof(renderMarkdownWithContext));
    }

    /// <summary>Creates a context-aware Markdown inline render extension registration.</summary>
    public static MarkdownInlineMarkdownRenderExtension CreateContextual(
        string name,
        Type inlineType,
        MarkdownInlineMarkdownContextualRenderer renderMarkdown) =>
        new MarkdownInlineMarkdownRenderExtension(
            name,
            inlineType,
            renderMarkdown == null
                ? throw new ArgumentNullException(nameof(renderMarkdown))
                : new MarkdownInlineMarkdownRenderer((inline, options) => renderMarkdown(inline, new MarkdownInlineMarkdownRenderContext(options, null))),
            renderMarkdown);

    /// <summary>Creates a Markdown inline render extension registration.</summary>
    public MarkdownInlineMarkdownRenderExtension(
        string name,
        Type inlineType,
        MarkdownInlineMarkdownRenderer renderMarkdown)
        : this(
            name,
            inlineType,
            renderMarkdown,
            renderMarkdown == null
                ? throw new ArgumentNullException(nameof(renderMarkdown))
                : new MarkdownInlineMarkdownContextualRenderer((inline, context) => renderMarkdown(inline, context.Options))) {
    }

    /// <summary>Stable extension name used for inspection or de-duplication.</summary>
    public string Name { get; }

    /// <summary>The inline type this extension handles.</summary>
    public Type InlineType { get; }

    /// <summary>Markdown rendering callback.</summary>
    public MarkdownInlineMarkdownRenderer RenderMarkdown { get; }

    /// <summary>Context-aware Markdown rendering callback.</summary>
    public MarkdownInlineMarkdownContextualRenderer RenderMarkdownWithContext { get; }

    /// <summary>Returns <see langword="true"/> when the extension can render the provided inline node.</summary>
    public bool Matches(IMarkdownInline inline) => inline != null && InlineType.IsInstanceOfType(inline);
}

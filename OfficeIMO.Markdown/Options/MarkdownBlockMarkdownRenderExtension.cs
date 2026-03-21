namespace OfficeIMO.Markdown;

/// <summary>
/// Delegate used to override markdown emitted for a specific block type during document serialization.
/// Returning <see langword="null"/> falls back to the block's built-in markdown rendering.
/// </summary>
public delegate string? MarkdownBlockMarkdownRenderer(IMarkdownBlock block, MarkdownWriteOptions options);

/// <summary>
/// Delegate used to override markdown emitted for a specific block type during document serialization with access
/// to the full markdown write context.
/// Returning <see langword="null"/> falls back to the block's built-in markdown rendering.
/// </summary>
public delegate string? MarkdownBlockMarkdownContextualRenderer(IMarkdownBlock block, MarkdownWriteContext context);

/// <summary>
/// Named markdown writer extension that can override emitted markdown for a specific block type.
/// </summary>
public sealed class MarkdownBlockMarkdownRenderExtension {
    private MarkdownBlockMarkdownRenderExtension(
        string name,
        Type blockType,
        MarkdownBlockMarkdownContextualRenderer renderMarkdown) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Extension name cannot be null or whitespace.", nameof(name));
        }

        Name = name.Trim();
        BlockType = blockType ?? throw new ArgumentNullException(nameof(blockType));
        RenderMarkdown = renderMarkdown ?? throw new ArgumentNullException(nameof(renderMarkdown));
    }

    /// <summary>Creates a context-aware markdown block render extension registration.</summary>
    public static MarkdownBlockMarkdownRenderExtension CreateContextual(
        string name,
        Type blockType,
        MarkdownBlockMarkdownContextualRenderer renderMarkdown) =>
        new MarkdownBlockMarkdownRenderExtension(name, blockType, renderMarkdown);

    /// <summary>Creates a markdown block render extension registration.</summary>
    public MarkdownBlockMarkdownRenderExtension(
        string name,
        Type blockType,
        MarkdownBlockMarkdownRenderer renderMarkdown)
        : this(
            name,
            blockType,
            renderMarkdown == null
                ? throw new ArgumentNullException(nameof(renderMarkdown))
                : new MarkdownBlockMarkdownContextualRenderer((block, context) => renderMarkdown(block, context.Options))) {
    }

    /// <summary>Stable extension name used for inspection or de-duplication.</summary>
    public string Name { get; }

    /// <summary>The block type this extension handles.</summary>
    public Type BlockType { get; }

    /// <summary>Markdown rendering callback.</summary>
    public MarkdownBlockMarkdownContextualRenderer RenderMarkdown { get; }

    /// <summary>Returns <see langword="true"/> when the extension can render the provided block.</summary>
    public bool Matches(IMarkdownBlock block) => block != null && BlockType.IsInstanceOfType(block);
}

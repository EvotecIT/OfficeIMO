namespace OfficeIMO.Markdown;

/// <summary>
/// Delegate used to override markdown emitted for a specific block type during document serialization.
/// Returning <see langword="null"/> falls back to the block's built-in markdown rendering.
/// </summary>
public delegate string? MarkdownBlockMarkdownRenderer(IMarkdownBlock block, MarkdownWriteOptions options);

/// <summary>
/// Named markdown writer extension that can override emitted markdown for a specific block type.
/// </summary>
public sealed class MarkdownBlockMarkdownRenderExtension {
    /// <summary>Creates a markdown block render extension registration.</summary>
    public MarkdownBlockMarkdownRenderExtension(
        string name,
        Type blockType,
        MarkdownBlockMarkdownRenderer renderMarkdown) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Extension name cannot be null or whitespace.", nameof(name));
        }

        Name = name.Trim();
        BlockType = blockType ?? throw new ArgumentNullException(nameof(blockType));
        RenderMarkdown = renderMarkdown ?? throw new ArgumentNullException(nameof(renderMarkdown));
    }

    /// <summary>Stable extension name used for inspection or de-duplication.</summary>
    public string Name { get; }

    /// <summary>The block type this extension handles.</summary>
    public Type BlockType { get; }

    /// <summary>Markdown rendering callback.</summary>
    public MarkdownBlockMarkdownRenderer RenderMarkdown { get; }

    /// <summary>Returns <see langword="true"/> when the extension can render the provided block.</summary>
    public bool Matches(IMarkdownBlock block) => block != null && BlockType.IsInstanceOfType(block);
}

namespace OfficeIMO.Markdown;

/// <summary>
/// Delegate used to override HTML emitted for a specific block type during HTML rendering.
/// Returning <see langword="null"/> falls back to the block's built-in HTML rendering.
/// </summary>
public delegate string? MarkdownBlockHtmlRenderer(IMarkdownBlock block, HtmlOptions options);

/// <summary>
/// Named HTML render extension that can override emitted HTML for a specific block type.
/// </summary>
public sealed class MarkdownBlockHtmlRenderExtension {
    /// <summary>Creates an HTML block render extension registration.</summary>
    public MarkdownBlockHtmlRenderExtension(
        string name,
        Type blockType,
        MarkdownBlockHtmlRenderer renderHtml) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Extension name cannot be null or whitespace.", nameof(name));
        }

        Name = name.Trim();
        BlockType = blockType ?? throw new ArgumentNullException(nameof(blockType));
        RenderHtml = renderHtml ?? throw new ArgumentNullException(nameof(renderHtml));
    }

    /// <summary>Stable extension name used for inspection or de-duplication.</summary>
    public string Name { get; }

    /// <summary>The block type this extension handles.</summary>
    public Type BlockType { get; }

    /// <summary>HTML rendering callback.</summary>
    public MarkdownBlockHtmlRenderer RenderHtml { get; }

    /// <summary>Returns <see langword="true"/> when the extension can render the provided block.</summary>
    public bool Matches(IMarkdownBlock block) => block != null && BlockType.IsInstanceOfType(block);
}

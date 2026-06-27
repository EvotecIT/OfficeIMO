namespace OfficeIMO.Markdown;

/// <summary>
/// Delegate used to override HTML emitted for a specific inline type during HTML rendering.
/// Returning <see langword="null"/> falls back to the inline's built-in HTML rendering.
/// </summary>
public delegate string? MarkdownInlineHtmlRenderer(IMarkdownInline inline, HtmlOptions options);

/// <summary>
/// Named HTML render extension that can override emitted HTML for a specific inline type.
/// </summary>
public sealed class MarkdownInlineHtmlRenderExtension {
    /// <summary>Creates an HTML inline render extension registration.</summary>
    public MarkdownInlineHtmlRenderExtension(
        string name,
        Type inlineType,
        MarkdownInlineHtmlRenderer renderHtml) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Extension name cannot be null or whitespace.", nameof(name));
        }

        Name = name.Trim();
        InlineType = inlineType ?? throw new ArgumentNullException(nameof(inlineType));
        RenderHtml = renderHtml ?? throw new ArgumentNullException(nameof(renderHtml));
    }

    /// <summary>Stable extension name used for inspection or de-duplication.</summary>
    public string Name { get; }

    /// <summary>The inline type this extension handles.</summary>
    public Type InlineType { get; }

    /// <summary>HTML rendering callback.</summary>
    public MarkdownInlineHtmlRenderer RenderHtml { get; }

    /// <summary>Returns <see langword="true"/> when the extension can render the provided inline node.</summary>
    public bool Matches(IMarkdownInline inline) => inline != null && InlineType.IsInstanceOfType(inline);
}

namespace OfficeIMO.Markdown;

/// <summary>
/// Delegate used to override HTML emitted for a specific inline type during HTML rendering.
/// Returning <see langword="null"/> falls back to the inline's built-in HTML rendering.
/// </summary>
public delegate string? MarkdownInlineHtmlRenderer(IMarkdownInline inline, HtmlOptions options);

/// <summary>
/// Delegate used to override HTML emitted for a specific inline type during HTML rendering with access
/// to the current inline render context.
/// Returning <see langword="null"/> falls back to the inline's built-in HTML rendering.
/// </summary>
public delegate string? MarkdownInlineHtmlContextualRenderer(IMarkdownInline inline, MarkdownInlineHtmlRenderContext context);

/// <summary>
/// Named HTML render extension that can override emitted HTML for a specific inline type.
/// </summary>
public sealed class MarkdownInlineHtmlRenderExtension {
    private MarkdownInlineHtmlRenderExtension(
        string name,
        Type inlineType,
        MarkdownInlineHtmlRenderer renderHtml,
        MarkdownInlineHtmlContextualRenderer renderHtmlWithContext) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Extension name cannot be null or whitespace.", nameof(name));
        }

        Name = name.Trim();
        InlineType = inlineType ?? throw new ArgumentNullException(nameof(inlineType));
        RenderHtml = renderHtml ?? throw new ArgumentNullException(nameof(renderHtml));
        RenderHtmlWithContext = renderHtmlWithContext ?? throw new ArgumentNullException(nameof(renderHtmlWithContext));
    }

    /// <summary>Creates a context-aware HTML inline render extension registration.</summary>
    public static MarkdownInlineHtmlRenderExtension CreateContextual(
        string name,
        Type inlineType,
        MarkdownInlineHtmlContextualRenderer renderHtml) =>
        new MarkdownInlineHtmlRenderExtension(
            name,
            inlineType,
            renderHtml == null
                ? throw new ArgumentNullException(nameof(renderHtml))
                : new MarkdownInlineHtmlRenderer((inline, options) => renderHtml(inline, new MarkdownInlineHtmlRenderContext(options, null))),
            renderHtml);

    /// <summary>Creates an HTML inline render extension registration.</summary>
    public MarkdownInlineHtmlRenderExtension(
        string name,
        Type inlineType,
        MarkdownInlineHtmlRenderer renderHtml)
        : this(
            name,
            inlineType,
            renderHtml,
            renderHtml == null
                ? throw new ArgumentNullException(nameof(renderHtml))
                : new MarkdownInlineHtmlContextualRenderer((inline, context) => renderHtml(inline, context.Options))) {
    }

    /// <summary>Stable extension name used for inspection or de-duplication.</summary>
    public string Name { get; }

    /// <summary>The inline type this extension handles.</summary>
    public Type InlineType { get; }

    /// <summary>HTML rendering callback.</summary>
    public MarkdownInlineHtmlRenderer RenderHtml { get; }

    /// <summary>Context-aware HTML rendering callback.</summary>
    public MarkdownInlineHtmlContextualRenderer RenderHtmlWithContext { get; }

    /// <summary>Returns <see langword="true"/> when the extension can render the provided inline node.</summary>
    public bool Matches(IMarkdownInline inline) => inline != null && InlineType.IsInstanceOfType(inline);
}

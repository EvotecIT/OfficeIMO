namespace OfficeIMO.Markdown;

/// <summary>
/// Delegate used to override HTML emitted for a block matched by its final syntax-tree node.
/// Returning <see langword="null"/> falls back to the block's built-in or type-based HTML rendering.
/// </summary>
public delegate string? MarkdownSyntaxBlockHtmlContextualRenderer(
    IMarkdownBlock block,
    MarkdownSyntaxNode syntaxNode,
    MarkdownBodyRenderContext context);

/// <summary>
/// Named HTML render extension that can override emitted HTML for blocks with a specific final syntax kind.
/// </summary>
public sealed class MarkdownSyntaxBlockHtmlRenderExtension {
    private MarkdownSyntaxBlockHtmlRenderExtension(
        string name,
        MarkdownSyntaxKind kind,
        string? customKind,
        MarkdownSyntaxBlockHtmlContextualRenderer renderHtml) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Extension name cannot be null or whitespace.", nameof(name));
        }

        Name = name.Trim();
        Kind = kind;
        CustomKind = NormalizeCustomKind(customKind);
        RenderHtml = renderHtml ?? throw new ArgumentNullException(nameof(renderHtml));
    }

    /// <summary>Creates a context-aware HTML block render extension registration matched by syntax kind.</summary>
    public static MarkdownSyntaxBlockHtmlRenderExtension CreateContextual(
        string name,
        MarkdownSyntaxKind kind,
        MarkdownSyntaxBlockHtmlContextualRenderer renderHtml,
        string? customKind = null) =>
        new MarkdownSyntaxBlockHtmlRenderExtension(name, kind, customKind, renderHtml);

    /// <summary>Stable extension name used for inspection or de-duplication.</summary>
    public string Name { get; }

    /// <summary>The final syntax kind this extension handles.</summary>
    public MarkdownSyntaxKind Kind { get; }

    /// <summary>Optional custom extension kind to require when matching <see cref="MarkdownSyntaxNode.CustomKind"/>.</summary>
    public string? CustomKind { get; }

    /// <summary>Context-aware HTML rendering callback.</summary>
    public MarkdownSyntaxBlockHtmlContextualRenderer RenderHtml { get; }

    /// <summary>Returns <see langword="true"/> when the extension can render the supplied syntax node.</summary>
    public bool Matches(MarkdownSyntaxNode syntaxNode) =>
        syntaxNode != null
        && syntaxNode.Kind == Kind
        && (CustomKind == null || string.Equals(CustomKind, syntaxNode.CustomKind, StringComparison.Ordinal));

    private static string? NormalizeCustomKind(string? customKind) =>
        string.IsNullOrWhiteSpace(customKind) ? null : customKind!.Trim();
}

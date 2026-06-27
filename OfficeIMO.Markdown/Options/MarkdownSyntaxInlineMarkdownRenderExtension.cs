namespace OfficeIMO.Markdown;

/// <summary>
/// Delegate used to override Markdown emitted for an inline matched by its final syntax-tree node.
/// Returning <see langword="null"/> falls back to the inline's built-in or type-based Markdown rendering.
/// </summary>
public delegate string? MarkdownSyntaxInlineMarkdownContextualRenderer(
    IMarkdownInline inline,
    MarkdownSyntaxNode syntaxNode,
    MarkdownInlineMarkdownRenderContext context);

/// <summary>
/// Named Markdown writer extension that can override emitted Markdown for inlines with a specific final syntax kind.
/// </summary>
public sealed class MarkdownSyntaxInlineMarkdownRenderExtension {
    private MarkdownSyntaxInlineMarkdownRenderExtension(
        string name,
        MarkdownSyntaxKind kind,
        string? customKind,
        MarkdownSyntaxInlineMarkdownContextualRenderer renderMarkdown) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Extension name cannot be null or whitespace.", nameof(name));
        }

        Name = name.Trim();
        Kind = kind;
        CustomKind = NormalizeCustomKind(customKind);
        RenderMarkdown = renderMarkdown ?? throw new ArgumentNullException(nameof(renderMarkdown));
    }

    /// <summary>Creates a context-aware Markdown inline render extension registration matched by syntax kind.</summary>
    public static MarkdownSyntaxInlineMarkdownRenderExtension CreateContextual(
        string name,
        MarkdownSyntaxKind kind,
        MarkdownSyntaxInlineMarkdownContextualRenderer renderMarkdown,
        string? customKind = null) =>
        new MarkdownSyntaxInlineMarkdownRenderExtension(name, kind, customKind, renderMarkdown);

    /// <summary>Stable extension name used for inspection or de-duplication.</summary>
    public string Name { get; }

    /// <summary>The final syntax kind this extension handles.</summary>
    public MarkdownSyntaxKind Kind { get; }

    /// <summary>Optional custom extension kind to require when matching <see cref="MarkdownSyntaxNode.CustomKind"/>.</summary>
    public string? CustomKind { get; }

    /// <summary>Context-aware Markdown rendering callback.</summary>
    public MarkdownSyntaxInlineMarkdownContextualRenderer RenderMarkdown { get; }

    /// <summary>Returns <see langword="true"/> when the extension can render the supplied syntax node.</summary>
    public bool Matches(MarkdownSyntaxNode syntaxNode) =>
        syntaxNode != null
        && syntaxNode.Kind == Kind
        && (CustomKind == null || string.Equals(CustomKind, syntaxNode.CustomKind, StringComparison.Ordinal));

    private static string? NormalizeCustomKind(string? customKind) =>
        string.IsNullOrWhiteSpace(customKind) ? null : customKind!.Trim();
}

namespace OfficeIMO.Markdown;

/// <summary>
/// Delegate used to override Markdown emitted for a block matched by its final syntax-tree node.
/// Returning <see langword="null"/> falls back to the block's built-in or type-based Markdown rendering.
/// </summary>
public delegate string? MarkdownSyntaxBlockMarkdownContextualRenderer(
    IMarkdownBlock block,
    MarkdownSyntaxNode syntaxNode,
    MarkdownWriteContext context);

/// <summary>
/// Named Markdown writer extension that can override emitted Markdown for blocks with a specific final syntax kind.
/// </summary>
public sealed class MarkdownSyntaxBlockMarkdownRenderExtension {
    private MarkdownSyntaxBlockMarkdownRenderExtension(
        string name,
        MarkdownSyntaxKind kind,
        string? customKind,
        MarkdownSyntaxBlockMarkdownContextualRenderer renderMarkdown) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Extension name cannot be null or whitespace.", nameof(name));
        }

        Name = name.Trim();
        Kind = kind;
        CustomKind = NormalizeCustomKind(customKind);
        RenderMarkdown = renderMarkdown ?? throw new ArgumentNullException(nameof(renderMarkdown));
    }

    /// <summary>Creates a context-aware Markdown block render extension registration matched by syntax kind.</summary>
    public static MarkdownSyntaxBlockMarkdownRenderExtension CreateContextual(
        string name,
        MarkdownSyntaxKind kind,
        MarkdownSyntaxBlockMarkdownContextualRenderer renderMarkdown,
        string? customKind = null) =>
        new MarkdownSyntaxBlockMarkdownRenderExtension(name, kind, customKind, renderMarkdown);

    /// <summary>Stable extension name used for inspection or de-duplication.</summary>
    public string Name { get; }

    /// <summary>The final syntax kind this extension handles.</summary>
    public MarkdownSyntaxKind Kind { get; }

    /// <summary>Optional custom extension kind to require when matching <see cref="MarkdownSyntaxNode.CustomKind"/>.</summary>
    public string? CustomKind { get; }

    /// <summary>Context-aware Markdown rendering callback.</summary>
    public MarkdownSyntaxBlockMarkdownContextualRenderer RenderMarkdown { get; }

    /// <summary>Returns <see langword="true"/> when the extension can render the supplied syntax node.</summary>
    public bool Matches(MarkdownSyntaxNode syntaxNode) =>
        syntaxNode != null
        && syntaxNode.Kind == Kind
        && (CustomKind == null || string.Equals(CustomKind, syntaxNode.CustomKind, StringComparison.Ordinal));

    private static string? NormalizeCustomKind(string? customKind) =>
        string.IsNullOrWhiteSpace(customKind) ? null : customKind!.Trim();
}

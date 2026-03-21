namespace OfficeIMO.Markdown;

/// <summary>
/// Extended block syntax contract that provides helper APIs for building child syntax nodes.
/// Prefer this over <see cref="ISyntaxMarkdownBlock"/> for custom block extensions that need
/// to compose nested block or inline syntax using the same rules as the core reader.
/// </summary>
public interface ISyntaxMarkdownBlockWithContext {
    /// <summary>
    /// Builds the syntax-tree node for this block using the supplied builder context.
    /// </summary>
    /// <param name="context">Helper context for building nested block and inline syntax nodes.</param>
    /// <param name="span">Source span mapped to the block when available.</param>
    /// <returns>A syntax node representing this block.</returns>
    MarkdownSyntaxNode BuildSyntaxNode(MarkdownBlockSyntaxBuilderContext context, MarkdownSourceSpan? span);
}

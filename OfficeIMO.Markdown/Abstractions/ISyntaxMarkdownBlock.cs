namespace OfficeIMO.Markdown;

/// <summary>
/// Allows a block node to control the syntax-tree node emitted by
/// <see cref="MarkdownReader.ParseWithSyntaxTree(string, MarkdownReaderOptions?)"/>.
/// Custom block parser and fenced-block extensions can implement this to contribute
/// precise AST nodes instead of falling back to generic <see cref="MarkdownSyntaxKind.Unknown"/> output.
/// </summary>
public interface ISyntaxMarkdownBlock {
    /// <summary>
    /// Builds the syntax-tree node for this block.
    /// </summary>
    /// <param name="span">Source span mapped to the block when available.</param>
    /// <returns>A syntax node representing this block.</returns>
    MarkdownSyntaxNode BuildSyntaxNode(MarkdownSourceSpan? span);
}

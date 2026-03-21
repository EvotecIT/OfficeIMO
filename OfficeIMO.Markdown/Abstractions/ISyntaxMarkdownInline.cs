namespace OfficeIMO.Markdown;

/// <summary>
/// Allows a custom inline node to control the syntax-tree node shape emitted by
/// <see cref="MarkdownReader.ParseWithSyntaxTree(string, MarkdownReaderOptions?)"/>.
/// </summary>
public interface ISyntaxMarkdownInline {
    /// <summary>
    /// Builds the syntax-tree node for this inline.
    /// </summary>
    /// <param name="context">Helper context for building nested inline syntax nodes.</param>
    /// <param name="span">Source span mapped to the inline when available.</param>
    /// <returns>A syntax node representing this inline.</returns>
    MarkdownSyntaxNode BuildSyntaxNode(MarkdownInlineSyntaxBuilderContext context, MarkdownSourceSpan? span);
}

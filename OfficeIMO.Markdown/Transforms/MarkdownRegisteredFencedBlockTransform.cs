namespace OfficeIMO.Markdown;

/// <summary>
/// Reapplies registered fenced-block factories to parsed fenced <see cref="CodeBlock"/> nodes
/// throughout the document tree, so nested hosts and fragment parsers converge on the same typed AST.
/// </summary>
public sealed class MarkdownRegisteredFencedBlockTransform : IMarkdownDocumentTransform {
    private readonly IReadOnlyList<MarkdownFencedBlockExtension> _extensions;

    /// <summary>
    /// Creates a transform that upgrades matching fenced <see cref="CodeBlock"/> nodes
    /// using the supplied registered fenced-block extensions.
    /// </summary>
    public MarkdownRegisteredFencedBlockTransform(IEnumerable<MarkdownFencedBlockExtension> extensions) {
        if (extensions == null) {
            throw new ArgumentNullException(nameof(extensions));
        }

        var registered = new List<MarkdownFencedBlockExtension>();
        foreach (var extension in extensions) {
            if (extension != null) {
                registered.Add(extension);
            }
        }

        _extensions = registered;
    }

    /// <inheritdoc />
    public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (_extensions.Count == 0) {
            return document;
        }

        MarkdownDocumentBlockRewriter.RewriteDocument(document, RewriteBlock);
        return document;
    }

    private IMarkdownBlock RewriteBlock(IMarkdownBlock block) {
        if (block is not CodeBlock codeBlock) {
            return block;
        }

        if (string.IsNullOrWhiteSpace(codeBlock.Language)) {
            return block;
        }

        return MarkdownReader.TryCreateExtendedFencedBlock(
                   _extensions,
                   codeBlock.InfoString,
                   codeBlock.Content,
                   codeBlock.IsFenced,
                   codeBlock.Caption)
               ?? block;
    }
}

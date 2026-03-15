namespace OfficeIMO.Markdown;

/// <summary>
/// Removes empty standalone <c>#</c> separator headings that appear immediately before real headings.
/// </summary>
/// <remarks>
/// This is intended for compatibility-oriented transcript/document cleanup where a model emitted
/// a stray single-hash line before the actual heading content. Ordinary empty headings that are not
/// directly followed by another heading are preserved.
/// </remarks>
/// <example>
/// <code>
/// var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
/// options.DocumentTransforms.Add(new MarkdownStandaloneHashHeadingSeparatorTransform());
///
/// var document = MarkdownReader.Parse("#\n## Result", options);
/// </code>
/// </example>
public sealed class MarkdownStandaloneHashHeadingSeparatorTransform : IMarkdownDocumentTransform {
    /// <inheritdoc />
    public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (context == null) {
            throw new ArgumentNullException(nameof(context));
        }

        MarkdownDocumentBlockListExpander.RewriteDocument(document, context, RewriteBlocks);
        return document;
    }

    private static List<IMarkdownBlock> RewriteBlocks(
        IReadOnlyList<IMarkdownBlock> blocks,
        MarkdownDocumentTransformContext context) {
        var rewritten = new List<IMarkdownBlock>(blocks.Count);
        for (var i = 0; i < blocks.Count; i++) {
            var block = blocks[i];
            if (block == null) {
                continue;
            }

            if (IsStandaloneHashSeparatorHeading(block)
                && TryFindNextNonEmptyBlock(blocks, i + 1, out var nextBlock)
                && nextBlock is HeadingBlock) {
                continue;
            }

            rewritten.Add(block);
        }

        return rewritten;
    }

    private static bool IsStandaloneHashSeparatorHeading(IMarkdownBlock block) {
        return block is HeadingBlock heading
               && heading.Level == 1
               && string.IsNullOrWhiteSpace(heading.Text);
    }

    private static bool TryFindNextNonEmptyBlock(
        IReadOnlyList<IMarkdownBlock> blocks,
        int startIndex,
        out IMarkdownBlock? nextBlock) {
        for (var i = startIndex; i < blocks.Count; i++) {
            nextBlock = blocks[i];
            if (nextBlock != null) {
                return true;
            }
        }

        nextBlock = null;
        return false;
    }
}

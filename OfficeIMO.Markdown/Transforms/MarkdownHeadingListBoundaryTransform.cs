namespace OfficeIMO.Markdown;

/// <summary>
/// Splits heading content when an unordered strong-label list marker was emitted directly after the heading text.
/// </summary>
/// <remarks>
/// This transform is intended for recoverable heading-level cleanup where markdown already parsed into a heading
/// block, but the heading text still contains an inline list marker that should begin a new unordered list item.
/// The split is performed directly on the parsed inline AST rather than by round-tripping the heading back through
/// markdown parsing.
/// </remarks>
/// <example>
/// <code>
/// var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
/// options.DocumentTransforms.Add(new MarkdownHeadingListBoundaryTransform());
///
/// var document = MarkdownReader.Parse("## Summary- **Item:** value", options);
/// </code>
/// </example>
public sealed class MarkdownHeadingListBoundaryTransform : IMarkdownDocumentTransform {
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

            if (block is HeadingBlock heading
                && TryRewriteHeading(heading, out var expandedBlocks)) {
                rewritten.AddRange(expandedBlocks);
                continue;
            }

            rewritten.Add(block);
        }

        return rewritten;
    }

    private static bool TryRewriteHeading(
        HeadingBlock heading,
        out IReadOnlyList<IMarkdownBlock> blocks) {
        if (!MarkdownInlineTransformHelpers.TrySplitHeadingListBoundary(
                heading.Inlines,
                out var head,
                out _,
                out var tail)) {
            blocks = Array.Empty<IMarkdownBlock>();
            return false;
        }

        head = MarkdownInlineTransformHelpers.TrimWhitespace(head, trimStart: true, trimEnd: true);
        tail = MarkdownInlineTransformHelpers.TrimWhitespace(tail, trimStart: true, trimEnd: true);
        if (!MarkdownInlineTransformHelpers.HasVisibleContent(head)
            || !MarkdownInlineTransformHelpers.StartsWithStrong(tail)) {
            blocks = Array.Empty<IMarkdownBlock>();
            return false;
        }

        var list = new UnorderedListBlock();
        foreach (var item in MarkdownInlineTransformHelpers.ExpandCompactStrongLabelListItems(tail, level: 0, forceLoose: false)) {
            list.Items.Add(item);
        }

        if (list.Items.Count == 0) {
            blocks = Array.Empty<IMarkdownBlock>();
            return false;
        }

        blocks = new IMarkdownBlock[] {
            new HeadingBlock(heading.Level, head),
            list
        };
        return true;
    }
}

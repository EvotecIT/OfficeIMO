namespace OfficeIMO.Markdown;

/// <summary>
/// Splits paragraph content when a prose label ending with a colon was emitted directly before a list marker.
/// </summary>
/// <remarks>
/// This transform is intended for recoverable paragraph-level cleanup where markdown already parsed into a
/// paragraph block, but the paragraph text still contains an inline list marker that should start a new list item.
/// The split is performed directly on the parsed inline AST rather than by reparsing rendered markdown.
/// </remarks>
/// <example>
/// <code>
/// var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
/// options.DocumentTransforms.Add(new MarkdownColonListBoundaryTransform());
///
/// var document = MarkdownReader.Parse("Next step:- **Item**", options);
/// </code>
/// </example>
public sealed class MarkdownColonListBoundaryTransform : IMarkdownDocumentTransform {
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

            if (block is ParagraphBlock paragraph
                && TryRewriteParagraph(paragraph, out var expandedBlocks)) {
                rewritten.AddRange(expandedBlocks);
                continue;
            }

            rewritten.Add(block);
        }

        return rewritten;
    }

    private static bool TryRewriteParagraph(
        ParagraphBlock paragraph,
        out IReadOnlyList<IMarkdownBlock> blocks) {
        if (!MarkdownInlineTransformHelpers.TrySplitColonListBoundary(
                paragraph.Inlines,
                out var head,
                out _,
                out var tail)) {
            blocks = Array.Empty<IMarkdownBlock>();
            return false;
        }

        head = MarkdownInlineTransformHelpers.TrimWhitespace(head, trimStart: false, trimEnd: true);
        tail = MarkdownInlineTransformHelpers.TrimWhitespace(tail, trimStart: true, trimEnd: true);
        if (!MarkdownInlineTransformHelpers.HasVisibleContent(head)
            || !MarkdownInlineTransformHelpers.HasVisibleContent(tail)) {
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
            new ParagraphBlock(head),
            list
        };
        return true;
    }
}

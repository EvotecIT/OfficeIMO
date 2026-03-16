namespace OfficeIMO.Markdown;

/// <summary>
/// Splits compact unordered strong-label list boundaries that were emitted inline after punctuation or symbols.
/// </summary>
/// <remarks>
/// This transform is intended for recoverable cleanup where markdown already parsed into a paragraph block or a
/// simple unordered list item, but the content still contains an inline list marker that should begin a new list item.
/// The split is performed directly on the parsed inline AST rather than by reparsing markdown text.
/// </remarks>
/// <example>
/// <code>
/// var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
/// options.DocumentTransforms.Add(new MarkdownCompactStrongLabelListBoundaryTransform());
///
/// var document = MarkdownReader.Parse("✅- **FSMO:** ok", options);
/// </code>
/// </example>
public sealed class MarkdownCompactStrongLabelListBoundaryTransform : IMarkdownDocumentTransform {
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

            switch (block) {
                case ParagraphBlock paragraph when TryRewriteParagraph(paragraph, out var expandedBlocks):
                    rewritten.AddRange(expandedBlocks);
                    break;
                case UnorderedListBlock unordered:
                    RewriteUnorderedList(unordered);
                    rewritten.Add(unordered);
                    break;
                default:
                    rewritten.Add(block);
                    break;
            }
        }

        return rewritten;
    }

    private static void RewriteUnorderedList(UnorderedListBlock list) {
        if (list.Items.Count == 0) {
            return;
        }

        var rewrittenItems = new List<ListItem>(list.Items.Count);
        for (var i = 0; i < list.Items.Count; i++) {
            var item = list.Items[i];
            if (item == null) {
                continue;
            }

            if (TryRewriteSimpleListItem(item, out var splitItems)) {
                rewrittenItems.AddRange(splitItems);
            } else {
                rewrittenItems.Add(item);
            }
        }

        list.Items.Clear();
        for (var i = 0; i < rewrittenItems.Count; i++) {
            list.Items.Add(rewrittenItems[i]);
        }
    }

    private static bool TryRewriteParagraph(
        ParagraphBlock paragraph,
        out IReadOnlyList<IMarkdownBlock> blocks) {
        if (!MarkdownInlineTransformHelpers.TrySplitCompactStrongLabelBoundary(
                paragraph.Inlines,
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

        blocks = new IMarkdownBlock[] {
            new ParagraphBlock(head),
            list
        };
        return list.Items.Count > 0;
    }

    private static bool TryRewriteSimpleListItem(
        ListItem item,
        out IReadOnlyList<ListItem> items) {
        if (item.IsTask || item.AdditionalParagraphs.Count > 0 || item.Children.Count > 0) {
            items = Array.Empty<ListItem>();
            return false;
        }

        if (!MarkdownInlineTransformHelpers.TrySplitCompactStrongLabelBoundary(
                item.Content,
                out var head,
                out _,
                out var tail)) {
            items = Array.Empty<ListItem>();
            return false;
        }

        head = MarkdownInlineTransformHelpers.TrimWhitespace(head, trimStart: true, trimEnd: true);
        tail = MarkdownInlineTransformHelpers.TrimWhitespace(tail, trimStart: true, trimEnd: true);
        if (!MarkdownInlineTransformHelpers.HasVisibleContent(head)
            || !MarkdownInlineTransformHelpers.StartsWithStrong(tail)) {
            items = Array.Empty<ListItem>();
            return false;
        }

        items = MarkdownInlineTransformHelpers.ExpandCompactStrongLabelListItems(item.Content, item.Level, item.ForceLoose);
        return items.Count > 1;
    }
}

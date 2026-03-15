namespace OfficeIMO.Markdown;

internal static class MarkdownDocumentBlockListExpander {
    public static void RewriteDocument(
        MarkdownDoc document,
        MarkdownDocumentTransformContext context,
        Func<IReadOnlyList<IMarkdownBlock>, MarkdownDocumentTransformContext, List<IMarkdownBlock>> blockListRewriter) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (context == null) {
            throw new ArgumentNullException(nameof(context));
        }

        if (blockListRewriter == null) {
            throw new ArgumentNullException(nameof(blockListRewriter));
        }

        document.ReplaceBlocks(RewriteBlocks(document.Blocks, context, blockListRewriter));
    }

    private static List<IMarkdownBlock> RewriteBlocks(
        IReadOnlyList<IMarkdownBlock> blocks,
        MarkdownDocumentTransformContext context,
        Func<IReadOnlyList<IMarkdownBlock>, MarkdownDocumentTransformContext, List<IMarkdownBlock>> blockListRewriter) {
        var rewritten = new List<IMarkdownBlock>(blocks.Count);
        for (var i = 0; i < blocks.Count; i++) {
            var block = blocks[i];
            if (block == null) {
                continue;
            }

            rewritten.Add(RewriteBlock(block, context, blockListRewriter));
        }

        return blockListRewriter(rewritten, context) ?? rewritten;
    }

    private static IMarkdownBlock RewriteBlock(
        IMarkdownBlock block,
        MarkdownDocumentTransformContext context,
        Func<IReadOnlyList<IMarkdownBlock>, MarkdownDocumentTransformContext, List<IMarkdownBlock>> blockListRewriter) {
        switch (block) {
            case QuoteBlock quote:
                RewriteMutableBlockList(quote.Children, context, blockListRewriter);
                return block;
            case DetailsBlock details:
                RewriteMutableBlockList(details.Children, context, blockListRewriter);
                return block;
            case OrderedListBlock ordered:
                RewriteListItems(ordered.Items, context, blockListRewriter);
                return block;
            case UnorderedListBlock unordered:
                RewriteListItems(unordered.Items, context, blockListRewriter);
                return block;
            case DefinitionListBlock definitions:
                RewriteDefinitionList(definitions, context, blockListRewriter);
                return block;
            case TableBlock table:
                RewriteTable(table, context, blockListRewriter);
                return block;
            case CalloutBlock callout:
                return RewriteCallout(callout, context, blockListRewriter);
            default:
                return block;
        }
    }

    private static void RewriteMutableBlockList(
        IList<IMarkdownBlock> blocks,
        MarkdownDocumentTransformContext context,
        Func<IReadOnlyList<IMarkdownBlock>, MarkdownDocumentTransformContext, List<IMarkdownBlock>> blockListRewriter) {
        if (blocks.Count == 0) {
            return;
        }

        var rewritten = RewriteBlocks(blocks.ToList(), context, blockListRewriter);
        blocks.Clear();
        for (var i = 0; i < rewritten.Count; i++) {
            blocks.Add(rewritten[i]);
        }
    }

    private static void RewriteListItems(
        IList<ListItem> items,
        MarkdownDocumentTransformContext context,
        Func<IReadOnlyList<IMarkdownBlock>, MarkdownDocumentTransformContext, List<IMarkdownBlock>> blockListRewriter) {
        for (var i = 0; i < items.Count; i++) {
            var item = items[i];
            if (item == null || item.Children.Count == 0) {
                continue;
            }

            RewriteMutableBlockList(item.Children, context, blockListRewriter);
        }
    }

    private static void RewriteDefinitionList(
        DefinitionListBlock block,
        MarkdownDocumentTransformContext context,
        Func<IReadOnlyList<IMarkdownBlock>, MarkdownDocumentTransformContext, List<IMarkdownBlock>> blockListRewriter) {
        var entries = block.Entries;
        for (var i = 0; i < entries.Count; i++) {
            var entry = entries[i];
            if (entry == null || entry.DefinitionBlocks.Count == 0) {
                continue;
            }

            RewriteMutableBlockList(entry.DefinitionBlocks, context, blockListRewriter);
        }
    }

    private static void RewriteTable(
        TableBlock table,
        MarkdownDocumentTransformContext context,
        Func<IReadOnlyList<IMarkdownBlock>, MarkdownDocumentTransformContext, List<IMarkdownBlock>> blockListRewriter) {
        RewriteTableCells(table.StructuredHeaders, context, blockListRewriter);
        if (table.StructuredRows == null) {
            return;
        }

        for (var i = 0; i < table.StructuredRows.Count; i++) {
            RewriteTableCells(table.StructuredRows[i], context, blockListRewriter);
        }
    }

    private static void RewriteTableCells(
        IEnumerable<TableCell>? cells,
        MarkdownDocumentTransformContext context,
        Func<IReadOnlyList<IMarkdownBlock>, MarkdownDocumentTransformContext, List<IMarkdownBlock>> blockListRewriter) {
        if (cells == null) {
            return;
        }

        foreach (var cell in cells) {
            if (cell == null || cell.Blocks.Count == 0) {
                continue;
            }

            RewriteMutableBlockList(cell.Blocks, context, blockListRewriter);
        }
    }

    private static CalloutBlock RewriteCallout(
        CalloutBlock block,
        MarkdownDocumentTransformContext context,
        Func<IReadOnlyList<IMarkdownBlock>, MarkdownDocumentTransformContext, List<IMarkdownBlock>> blockListRewriter) {
        if (block.ChildBlocks.Count == 0) {
            return block;
        }

        var rewrittenChildren = RewriteBlocks(block.ChildBlocks, context, blockListRewriter);
        return new CalloutBlock(block.Kind, block.TitleInlines, rewrittenChildren, block.SyntaxChildren);
    }
}

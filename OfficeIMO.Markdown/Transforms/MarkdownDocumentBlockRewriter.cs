namespace OfficeIMO.Markdown;

internal static class MarkdownDocumentBlockRewriter {
    public static void RewriteDocument(MarkdownDoc document, Func<IMarkdownBlock, IMarkdownBlock> blockRewriter) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (blockRewriter == null) {
            throw new ArgumentNullException(nameof(blockRewriter));
        }

        var rewritten = RewriteBlocks(document.Blocks, blockRewriter);
        document.ReplaceBlocks(rewritten);
    }

    private static List<IMarkdownBlock> RewriteBlocks(
        IEnumerable<IMarkdownBlock> blocks,
        Func<IMarkdownBlock, IMarkdownBlock> blockRewriter) {
        var rewritten = new List<IMarkdownBlock>();
        foreach (var block in blocks) {
            if (block == null) {
                continue;
            }

            rewritten.Add(RewriteBlock(block, blockRewriter));
        }

        return rewritten;
    }

    private static IMarkdownBlock RewriteBlock(
        IMarkdownBlock block,
        Func<IMarkdownBlock, IMarkdownBlock> blockRewriter) {
        IMarkdownBlock current = block;

        switch (current) {
            case QuoteBlock quote:
                RewriteMutableBlockList(quote.Children, blockRewriter);
                break;
            case DetailsBlock details:
                RewriteMutableBlockList(details.Children, blockRewriter);
                break;
            case OrderedListBlock ordered:
                RewriteListItems(ordered.Items, blockRewriter);
                break;
            case UnorderedListBlock unordered:
                RewriteListItems(unordered.Items, blockRewriter);
                break;
            case DefinitionListBlock definitions:
                RewriteDefinitionList(definitions, blockRewriter);
                break;
            case TableBlock table:
                RewriteTable(table, blockRewriter);
                break;
            case CalloutBlock callout:
                current = RewriteCallout(callout, blockRewriter);
                break;
        }

        return blockRewriter(current) ?? current;
    }

    private static void RewriteMutableBlockList(
        IList<IMarkdownBlock> blocks,
        Func<IMarkdownBlock, IMarkdownBlock> blockRewriter) {
        for (var i = 0; i < blocks.Count; i++) {
            var block = blocks[i];
            if (block == null) {
                continue;
            }

            blocks[i] = RewriteBlock(block, blockRewriter);
        }
    }

    private static void RewriteListItems(
        IList<ListItem> items,
        Func<IMarkdownBlock, IMarkdownBlock> blockRewriter) {
        for (var i = 0; i < items.Count; i++) {
            var item = items[i];
            if (item == null) {
                continue;
            }

            RewriteMutableBlockList(item.Children, blockRewriter);
        }
    }

    private static void RewriteDefinitionList(
        DefinitionListBlock block,
        Func<IMarkdownBlock, IMarkdownBlock> blockRewriter) {
        var entries = block.Entries;
        for (var i = 0; i < entries.Count; i++) {
            var entry = entries[i];
            if (entry == null) {
                continue;
            }

            RewriteMutableBlockList(entry.DefinitionBlocks, blockRewriter);
        }
    }

    private static void RewriteTable(
        TableBlock table,
        Func<IMarkdownBlock, IMarkdownBlock> blockRewriter) {
        RewriteTableCells(table.StructuredHeaders, blockRewriter);
        if (table.StructuredRows == null) {
            return;
        }

        for (var i = 0; i < table.StructuredRows.Count; i++) {
            RewriteTableCells(table.StructuredRows[i], blockRewriter);
        }
    }

    private static void RewriteTableCells(
        IEnumerable<TableCell>? cells,
        Func<IMarkdownBlock, IMarkdownBlock> blockRewriter) {
        if (cells == null) {
            return;
        }

        foreach (var cell in cells) {
            if (cell == null) {
                continue;
            }

            RewriteMutableBlockList(cell.Blocks, blockRewriter);
        }
    }

    private static IMarkdownBlock RewriteCallout(
        CalloutBlock block,
        Func<IMarkdownBlock, IMarkdownBlock> blockRewriter) {
        if (block.ChildBlocks.Count == 0) {
            return block;
        }

        var rewrittenChildren = RewriteBlocks(block.ChildBlocks, blockRewriter);
        return new CalloutBlock(block.Kind, block.TitleInlines, rewrittenChildren, block.SyntaxChildren);
    }
}

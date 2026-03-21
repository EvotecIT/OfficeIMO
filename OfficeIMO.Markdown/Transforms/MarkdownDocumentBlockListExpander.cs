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

            rewritten.Add(PreserveSourceSpan(block, RewriteBlock(block, context, blockListRewriter)));
        }

        var expanded = blockListRewriter(rewritten, context) ?? rewritten;
        ApplyChangedRangeSourceSpans(rewritten, expanded);
        return expanded;
    }

    private static IMarkdownBlock RewriteBlock(
        IMarkdownBlock block,
        MarkdownDocumentTransformContext context,
        Func<IReadOnlyList<IMarkdownBlock>, MarkdownDocumentTransformContext, List<IMarkdownBlock>> blockListRewriter) {
        switch (block) {
            case QuoteBlock quote:
                RewriteMutableBlockList(quote.Children, context, blockListRewriter);
                quote.ClearSyntaxCache();
                return block;
            case DetailsBlock details:
                RewriteMutableBlockList(details.Children, context, blockListRewriter);
                details.ClearSyntaxCache();
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
            case FootnoteDefinitionBlock footnote:
                return RewriteFootnote(footnote, context, blockListRewriter);
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
        var groups = block.Groups;
        for (var groupIndex = 0; groupIndex < groups.Count; groupIndex++) {
            var group = groups[groupIndex];
            if (group == null) {
                continue;
            }

            for (var definitionIndex = 0; definitionIndex < group.Definitions.Count; definitionIndex++) {
                var definition = group.Definitions[definitionIndex];
                if (definition == null || definition.Blocks.Count == 0) {
                    continue;
                }

                RewriteMutableBlockList(definition.Blocks, context, blockListRewriter);
            }
        }

        block.ClearSyntaxCache();
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

    private static FootnoteDefinitionBlock RewriteFootnote(
        FootnoteDefinitionBlock block,
        MarkdownDocumentTransformContext context,
        Func<IReadOnlyList<IMarkdownBlock>, MarkdownDocumentTransformContext, List<IMarkdownBlock>> blockListRewriter) {
        if (block.Blocks.Count == 0) {
            return block;
        }

        var rewrittenBlocks = RewriteBlocks(block.Blocks, context, blockListRewriter);
        return new FootnoteDefinitionBlock(block.Label, block.Text, rewrittenBlocks, syntaxChildren: null);
    }

    private static CalloutBlock RewriteCallout(
        CalloutBlock block,
        MarkdownDocumentTransformContext context,
        Func<IReadOnlyList<IMarkdownBlock>, MarkdownDocumentTransformContext, List<IMarkdownBlock>> blockListRewriter) {
        if (block.ChildBlocks.Count == 0) {
            return block;
        }

        var rewrittenChildren = RewriteBlocks(block.ChildBlocks, context, blockListRewriter);
        return new CalloutBlock(block.Kind, block.TitleInlines, rewrittenChildren, syntaxChildren: null);
    }

    private static IMarkdownBlock PreserveSourceSpan(IMarkdownBlock original, IMarkdownBlock rewritten) {
        if (ReferenceEquals(original, rewritten)
            || original is not MarkdownObject originalObject
            || rewritten is not MarkdownObject rewrittenObject
            || rewrittenObject.SourceSpan.HasValue
            || !originalObject.SourceSpan.HasValue) {
            return rewritten;
        }

        rewrittenObject.SourceSpan = originalObject.SourceSpan;
        return rewritten;
    }

    private static void ApplyChangedRangeSourceSpans(
        IReadOnlyList<IMarkdownBlock> before,
        IReadOnlyList<IMarkdownBlock> after) {
        if (before == null || after == null || before.Count == 0 || after.Count == 0) {
            return;
        }

        var change = MarkdownTransformSourceSpanHelper.ComputeChangedRange(
            MarkdownTransformSourceSpanHelper.CreateBlockFingerprints(before),
            MarkdownTransformSourceSpanHelper.CreateBlockFingerprints(after));
        if (change.CountAfter == 0) {
            return;
        }

        var affectedSourceSpan = MarkdownTransformSourceSpanHelper.AggregateBlockSpans(before, change.StartBefore, change.CountBefore);
        if (!affectedSourceSpan.HasValue) {
            return;
        }

        MarkdownTransformSourceSpanHelper.ApplyAffectedSpanToChangedBlocks(after, change, affectedSourceSpan);
    }
}

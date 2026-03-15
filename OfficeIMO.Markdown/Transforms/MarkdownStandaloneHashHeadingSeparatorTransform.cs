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

        document.ReplaceBlocks(RewriteBlocks(document.Blocks, context));
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

            rewritten.Add(RewriteBlock(block, context));
        }

        return rewritten;
    }

    private static IMarkdownBlock RewriteBlock(
        IMarkdownBlock block,
        MarkdownDocumentTransformContext context) {
        switch (block) {
            case QuoteBlock quote:
                RewriteMutableBlockList(quote.Children, context);
                return block;
            case DetailsBlock details:
                RewriteMutableBlockList(details.Children, context);
                return block;
            case OrderedListBlock ordered:
                RewriteListItems(ordered.Items, context);
                return block;
            case UnorderedListBlock unordered:
                RewriteListItems(unordered.Items, context);
                return block;
            case TableBlock table:
                RewriteTable(table, context);
                return block;
            case CalloutBlock callout:
                return RewriteCallout(callout, context);
            case DefinitionListBlock definitions:
                RewriteDefinitionList(definitions, context);
                return block;
            default:
                return block;
        }
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

    private static void RewriteMutableBlockList(
        IList<IMarkdownBlock> blocks,
        MarkdownDocumentTransformContext context) {
        if (blocks.Count == 0) {
            return;
        }

        var rewritten = RewriteBlocks(blocks.ToList(), context);
        blocks.Clear();
        for (var i = 0; i < rewritten.Count; i++) {
            blocks.Add(rewritten[i]);
        }
    }

    private static void RewriteListItems(
        IList<ListItem> items,
        MarkdownDocumentTransformContext context) {
        for (var i = 0; i < items.Count; i++) {
            var item = items[i];
            if (item == null || item.Children.Count == 0) {
                continue;
            }

            RewriteMutableBlockList(item.Children, context);
        }
    }

    private static void RewriteDefinitionList(
        DefinitionListBlock block,
        MarkdownDocumentTransformContext context) {
        var entries = block.Entries;
        for (var i = 0; i < entries.Count; i++) {
            var entry = entries[i];
            if (entry == null || entry.DefinitionBlocks.Count == 0) {
                continue;
            }

            RewriteMutableBlockList(entry.DefinitionBlocks, context);
        }
    }

    private static void RewriteTable(
        TableBlock table,
        MarkdownDocumentTransformContext context) {
        RewriteTableCells(table.StructuredHeaders, context);
        if (table.StructuredRows == null) {
            return;
        }

        for (var i = 0; i < table.StructuredRows.Count; i++) {
            RewriteTableCells(table.StructuredRows[i], context);
        }
    }

    private static void RewriteTableCells(
        IEnumerable<TableCell>? cells,
        MarkdownDocumentTransformContext context) {
        if (cells == null) {
            return;
        }

        foreach (var cell in cells) {
            if (cell == null || cell.Blocks.Count == 0) {
                continue;
            }

            RewriteMutableBlockList(cell.Blocks, context);
        }
    }

    private static CalloutBlock RewriteCallout(
        CalloutBlock block,
        MarkdownDocumentTransformContext context) {
        if (block.ChildBlocks.Count == 0) {
            return block;
        }

        var rewrittenChildren = RewriteBlocks(block.ChildBlocks, context);
        return new CalloutBlock(block.Kind, block.TitleInlines, rewrittenChildren, block.SyntaxChildren);
    }
}

namespace OfficeIMO.Markdown;

/// <summary>
/// Expands simple parsed definition-list entries into ordinary paragraphs.
/// </summary>
/// <remarks>
/// Use this for compatibility-oriented hosts that prefer narrative paragraph output over
/// grouped definition-list rendering, while still preserving complex definition-list entries.
/// Entries are converted only when the term is non-empty and the definition body is a single paragraph.
/// </remarks>
/// <example>
/// <code>
/// var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
/// options.DocumentTransforms.Add(new MarkdownSimpleDefinitionListParagraphTransform());
///
/// var document = MarkdownReader.Parse("Status: healthy\nImpact: none", options);
/// </code>
/// </example>
public sealed class MarkdownSimpleDefinitionListParagraphTransform : IMarkdownDocumentTransform {
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

            rewritten.AddRange(RewriteBlock(block, context));
        }

        return rewritten;
    }

    private static IReadOnlyList<IMarkdownBlock> RewriteBlock(
        IMarkdownBlock block,
        MarkdownDocumentTransformContext context) {
        switch (block) {
            case DefinitionListBlock definitions:
                return ExpandDefinitionList(definitions, context);
            case QuoteBlock quote:
                RewriteMutableBlockList(quote.Children, context);
                return new[] { block };
            case DetailsBlock details:
                RewriteMutableBlockList(details.Children, context);
                return new[] { block };
            case OrderedListBlock ordered:
                RewriteListItems(ordered.Items, context);
                return new[] { block };
            case UnorderedListBlock unordered:
                RewriteListItems(unordered.Items, context);
                return new[] { block };
            case TableBlock table:
                RewriteTable(table, context);
                return new[] { block };
            case CalloutBlock callout:
                return new[] { RewriteCallout(callout, context) };
            default:
                return new[] { block };
        }
    }

    private static IReadOnlyList<IMarkdownBlock> ExpandDefinitionList(
        DefinitionListBlock block,
        MarkdownDocumentTransformContext context) {
        if (block.Entries.Count == 0) {
            return new[] { (IMarkdownBlock)block };
        }

        var rewritten = new List<IMarkdownBlock>();
        DefinitionListBlock? pendingDefinitionList = null;

        for (var i = 0; i < block.Entries.Count; i++) {
            var entry = block.Entries[i];
            if (TryConvertEntryToParagraph(entry, block, context, out var paragraph)) {
                FlushPendingDefinitionList(rewritten, ref pendingDefinitionList);
                rewritten.Add(paragraph);
                continue;
            }

            pendingDefinitionList ??= new DefinitionListBlock();
            pendingDefinitionList.AddEntry(new DefinitionListEntry(
                entry.Term,
                RewriteBlocks(entry.DefinitionBlocks, context)));
        }

        FlushPendingDefinitionList(rewritten, ref pendingDefinitionList);
        return rewritten.Count == 0 ? new[] { (IMarkdownBlock)block } : rewritten;
    }

    private static bool TryConvertEntryToParagraph(
        DefinitionListEntry entry,
        DefinitionListBlock owner,
        MarkdownDocumentTransformContext context,
        out ParagraphBlock paragraph) {
        paragraph = null!;
        if (entry == null || entry.DefinitionBlocks.Count != 1 || entry.DefinitionBlocks[0] is not ParagraphBlock definitionParagraph) {
            return false;
        }

        var termMarkdown = entry.TermMarkdown?.Trim();
        var definitionMarkdown = definitionParagraph.Inlines.RenderMarkdown().Trim();
        if (string.IsNullOrEmpty(termMarkdown) || string.IsNullOrEmpty(definitionMarkdown)) {
            return false;
        }

        var readerOptions = owner.ReaderOptions
            ?? context.ReaderOptions
            ?? new MarkdownReaderOptions();
        var readerState = owner.ReaderState;
        var combined = termMarkdown + ": " + definitionMarkdown;
        paragraph = new ParagraphBlock(MarkdownReader.ParseInlineText(combined, readerOptions, readerState));
        return true;
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

    private static void FlushPendingDefinitionList(
        ICollection<IMarkdownBlock> blocks,
        ref DefinitionListBlock? pendingDefinitionList) {
        if (pendingDefinitionList == null || pendingDefinitionList.Entries.Count == 0) {
            pendingDefinitionList = null;
            return;
        }

        blocks.Add(pendingDefinitionList);
        pendingDefinitionList = null;
    }
}

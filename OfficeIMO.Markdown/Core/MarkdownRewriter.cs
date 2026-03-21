using System.Linq;

namespace OfficeIMO.Markdown;

#pragma warning disable CS1591

/// <summary>
/// Compatibility-first recursive rewriter for markdown block trees.
/// </summary>
public abstract class MarkdownRewriter {
    /// <summary>Rewrites the document in place and returns it.</summary>
    public virtual MarkdownDoc Rewrite(MarkdownDoc document) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        var rewritten = RewriteBlocks(document.Blocks);
        document.ReplaceBlocks(rewritten);
        return document;
    }

    /// <summary>Rewrites a block sequence in order.</summary>
    protected virtual List<IMarkdownBlock> RewriteBlocks(IEnumerable<IMarkdownBlock> blocks) {
        var rewritten = new List<IMarkdownBlock>();
        foreach (var block in blocks) {
            if (block == null) {
                continue;
            }

            rewritten.Add(RewriteBlock(block));
        }

        return rewritten;
    }

    /// <summary>Rewrites a single block after recursively rewriting its owned child blocks.</summary>
    protected virtual IMarkdownBlock RewriteBlock(IMarkdownBlock block) {
        IMarkdownBlock current = block ?? throw new ArgumentNullException(nameof(block));

        switch (current) {
            case QuoteBlock quote:
                RewriteMutableBlockList(quote.Children);
                quote.ClearSyntaxCache();
                break;
            case DetailsBlock details:
                RewriteMutableBlockList(details.Children);
                details.ClearSyntaxCache();
                break;
            case OrderedListBlock ordered:
                RewriteListItems(ordered.Items);
                break;
            case UnorderedListBlock unordered:
                RewriteListItems(unordered.Items);
                break;
            case DefinitionListBlock definitions:
                RewriteDefinitionList(definitions);
                break;
            case FootnoteDefinitionBlock footnote:
                current = RewriteFootnote(footnote);
                break;
            case TableBlock table:
                RewriteTable(table);
                break;
            case CalloutBlock callout:
                current = RewriteCallout(callout);
                break;
        }

        var rewritten = RewriteCurrentBlock(current) ?? current;
        return PreserveSourceSpan(current, rewritten);
    }

    /// <summary>Hook invoked after child blocks have already been rewritten.</summary>
    protected virtual IMarkdownBlock RewriteCurrentBlock(IMarkdownBlock block) => block;

    protected void RewriteMutableBlockList(IList<IMarkdownBlock> blocks) {
        for (var i = 0; i < blocks.Count; i++) {
            var block = blocks[i];
            if (block == null) {
                continue;
            }

            blocks[i] = RewriteBlock(block);
        }
    }

    protected virtual void RewriteListItems(IList<ListItem> items) {
        for (var i = 0; i < items.Count; i++) {
            var item = items[i];
            if (item == null) {
                continue;
            }

            var rewrittenBlocks = RewriteBlocks(item.BlockChildren);
            IReadOnlyList<IMarkdownInline> leadParagraphNodes = Array.Empty<IMarkdownInline>();
            if (rewrittenBlocks.Count > 0 && rewrittenBlocks[0] is ParagraphBlock leadParagraph) {
                leadParagraphNodes = leadParagraph.Inlines.Nodes
                    .Where(node => node != null)
                    .ToArray();
            }

            item.SyntaxChildren.Clear();
            item.Content.ReplaceItems(Array.Empty<IMarkdownInline>());
            item.AdditionalParagraphs.Clear();
            item.Children.Clear();

            var blockIndex = 0;
            if (leadParagraphNodes.Count > 0 || (rewrittenBlocks.Count > 0 && rewrittenBlocks[0] is ParagraphBlock)) {
                item.Content.ReplaceItems(leadParagraphNodes);
                blockIndex = 1;
            }

            while (blockIndex < rewrittenBlocks.Count && rewrittenBlocks[blockIndex] is ParagraphBlock additionalParagraph) {
                item.AdditionalParagraphs.Add(additionalParagraph.Inlines);
                blockIndex++;
            }

            for (; blockIndex < rewrittenBlocks.Count; blockIndex++) {
                item.Children.Add(rewrittenBlocks[blockIndex]);
            }
        }
    }

    protected virtual void RewriteDefinitionList(DefinitionListBlock block) {
        var groups = block.Groups;
        for (var groupIndex = 0; groupIndex < groups.Count; groupIndex++) {
            var group = groups[groupIndex];
            if (group == null) {
                continue;
            }

            for (var definitionIndex = 0; definitionIndex < group.Definitions.Count; definitionIndex++) {
                RewriteMutableBlockList(group.Definitions[definitionIndex].Blocks);
            }
        }

        block.ClearSyntaxCache();
    }

    protected virtual void RewriteTable(TableBlock table) {
        RewriteTableCells(table.StructuredHeaders);
        if (table.StructuredRows == null) {
            return;
        }

        for (var i = 0; i < table.StructuredRows.Count; i++) {
            RewriteTableCells(table.StructuredRows[i]);
        }
    }

    protected void RewriteTableCells(IEnumerable<TableCell>? cells) {
        if (cells == null) {
            return;
        }

        foreach (var cell in cells) {
            if (cell == null) {
                continue;
            }

            RewriteMutableBlockList(cell.Blocks);
        }
    }

    protected virtual FootnoteDefinitionBlock RewriteFootnote(FootnoteDefinitionBlock block) {
        if (block.Blocks.Count == 0) {
            return block;
        }

        var rewrittenBlocks = RewriteBlocks(block.Blocks);
        return new FootnoteDefinitionBlock(block.Label, block.Text, rewrittenBlocks, syntaxChildren: null);
    }

    protected virtual CalloutBlock RewriteCallout(CalloutBlock block) {
        if (block.ChildBlocks.Count == 0) {
            return block;
        }

        var rewrittenChildren = RewriteBlocks(block.ChildBlocks);
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
}

#pragma warning restore CS1591

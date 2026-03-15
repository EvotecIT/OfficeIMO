namespace OfficeIMO.Markdown;

/// <summary>
/// Applies AST-safe inline normalization to an existing <see cref="MarkdownDoc"/>.
/// </summary>
/// <remarks>
/// Use this when markdown or HTML has already been parsed into a document and you want to apply
/// inline-safe cleanup such as escaped code-span repair, tight strong-boundary spacing, tight colon spacing,
/// or tight parenthetical spacing without going back through the text pre-parse path.
/// </remarks>
public sealed class MarkdownInlineNormalizationTransform : IMarkdownDocumentTransform {
    /// <summary>
    /// Creates a document transform that applies AST-safe inline normalization.
    /// </summary>
    /// <param name="options">Inline normalization options. Only AST-safe options take effect.</param>
    public MarkdownInlineNormalizationTransform(MarkdownInputNormalizationOptions? options = null) {
        Options = options ?? new MarkdownInputNormalizationOptions();
    }

    /// <summary>
    /// Inline normalization options used by the transform.
    /// </summary>
    public MarkdownInputNormalizationOptions Options { get; }

    /// <inheritdoc />
    public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (context == null) {
            throw new ArgumentNullException(nameof(context));
        }

        var rewritten = RewriteBlockList(document.Blocks, out _);
        document.ReplaceBlocks(rewritten);
        return document;
    }

    private void NormalizeMutableBlockList(IList<IMarkdownBlock> blocks) {
        for (int i = 0; i < blocks.Count; i++) {
            var block = blocks[i];
            if (block == null) {
                continue;
            }

            blocks[i] = NormalizeBlock(block);
        }
    }

    private List<IMarkdownBlock> RewriteBlockList(IReadOnlyList<IMarkdownBlock> blocks, out bool changed) {
        changed = false;
        var rewritten = new List<IMarkdownBlock>(blocks.Count);
        for (int i = 0; i < blocks.Count; i++) {
            var block = blocks[i];
            if (block == null) {
                continue;
            }

            var normalized = NormalizeBlock(block);
            rewritten.Add(normalized);
            changed |= !ReferenceEquals(block, normalized);
        }

        return rewritten;
    }

    private IMarkdownBlock NormalizeBlock(IMarkdownBlock block) {
        switch (block) {
            case ParagraphBlock paragraph:
                MarkdownReader.NormalizeInlineSequenceInPlace(paragraph.Inlines, Options);
                return paragraph;
            case HeadingBlock heading:
                return MarkdownReader.NormalizeInlineSequenceInPlace(heading.Inlines, Options)
                    ? new HeadingBlock(heading.Level, heading.Inlines)
                    : heading;
            case SummaryBlock summary:
                MarkdownReader.NormalizeInlineSequenceInPlace(summary.Inlines, Options);
                return summary;
            case CalloutBlock callout:
                return NormalizeCallout(callout);
            case QuoteBlock quote:
                NormalizeMutableBlockList(quote.Children);
                return quote;
            case DetailsBlock details:
                if (details.Summary != null) {
                    NormalizeBlock(details.Summary);
                }
                NormalizeMutableBlockList(details.Children);
                return details;
            case OrderedListBlock ordered:
                NormalizeListItems(ordered.Items);
                return ordered;
            case UnorderedListBlock unordered:
                NormalizeListItems(unordered.Items);
                return unordered;
            case DefinitionListBlock definitions:
                NormalizeDefinitionList(definitions);
                return definitions;
            case FootnoteDefinitionBlock footnote:
                return NormalizeFootnote(footnote);
            case TableBlock table:
                NormalizeTable(table);
                return table;
            default:
                return block;
        }
    }

    private IMarkdownBlock NormalizeCallout(CalloutBlock callout) {
        bool titleChanged = MarkdownReader.NormalizeInlineSequenceInPlace(callout.TitleInlines, Options);
        var rewrittenChildren = RewriteChildBlocks(callout.ChildBlocks, out bool childChanged);
        if (!titleChanged && !childChanged) {
            return callout;
        }

        return new CalloutBlock(callout.Kind, callout.TitleInlines, rewrittenChildren, callout.SyntaxChildren);
    }

    private void NormalizeListItems(IList<ListItem> items) {
        for (int i = 0; i < items.Count; i++) {
            var item = items[i];
            if (item == null) {
                continue;
            }

            MarkdownReader.NormalizeInlineSequenceInPlace(item.Content, Options);
            for (int paragraphIndex = 0; paragraphIndex < item.AdditionalParagraphs.Count; paragraphIndex++) {
                MarkdownReader.NormalizeInlineSequenceInPlace(item.AdditionalParagraphs[paragraphIndex], Options);
            }

            NormalizeMutableBlockList(item.Children);
        }
    }

    private void NormalizeDefinitionList(DefinitionListBlock definitions) {
        for (int i = 0; i < definitions.Entries.Count; i++) {
            var entry = definitions.Entries[i];
            if (entry == null) {
                continue;
            }

            MarkdownReader.NormalizeInlineSequenceInPlace(entry.Term, Options);
            NormalizeMutableBlockList(entry.DefinitionBlocks);
        }
    }

    private IMarkdownBlock NormalizeFootnote(FootnoteDefinitionBlock footnote) {
        if (footnote.ParagraphBlocks.Count == 0) {
            return footnote;
        }

        var rewrittenParagraphs = new List<ParagraphBlock>(footnote.ParagraphBlocks.Count);
        bool changed = false;

        for (int i = 0; i < footnote.ParagraphBlocks.Count; i++) {
            var paragraph = footnote.ParagraphBlocks[i] ?? new ParagraphBlock(new InlineSequence());
            bool paragraphChanged = MarkdownReader.NormalizeInlineSequenceInPlace(paragraph.Inlines, Options);
            rewrittenParagraphs.Add(paragraph);
            changed |= paragraphChanged;
        }

        return changed
            ? new FootnoteDefinitionBlock(footnote.Label, footnote.Text, rewrittenParagraphs, footnote.SyntaxChildren)
            : footnote;
    }

    private void NormalizeTable(TableBlock table) {
        if (table.StructuredHeaders != null) {
            NormalizeTableCells(table.StructuredHeaders);
        }

        if (table.StructuredRows != null) {
            for (int i = 0; i < table.StructuredRows.Count; i++) {
                NormalizeTableCells(table.StructuredRows[i]);
            }
        }

        if (table.ParsedHeaders != null) {
            NormalizeInlineSequences(table.ParsedHeaders);
        }

        if (table.ParsedRows != null) {
            for (int i = 0; i < table.ParsedRows.Count; i++) {
                NormalizeInlineSequences(table.ParsedRows[i]);
            }
        }
    }

    private void NormalizeTableCells(IEnumerable<TableCell> cells) {
        foreach (var cell in cells) {
            if (cell == null) {
                continue;
            }

            NormalizeMutableBlockList(cell.Blocks);
        }
    }

    private void NormalizeInlineSequences(IEnumerable<InlineSequence> sequences) {
        foreach (var sequence in sequences) {
            MarkdownReader.NormalizeInlineSequenceInPlace(sequence, Options);
        }
    }

    private List<IMarkdownBlock> RewriteChildBlocks(IReadOnlyList<IMarkdownBlock> blocks, out bool changed) {
        return RewriteBlockList(blocks, out changed);
    }
}

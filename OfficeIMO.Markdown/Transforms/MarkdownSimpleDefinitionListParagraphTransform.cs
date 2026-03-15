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

            rewritten.AddRange(RewriteBlock(block, context));
        }

        return rewritten;
    }

    private static IReadOnlyList<IMarkdownBlock> RewriteBlock(
        IMarkdownBlock block,
        MarkdownDocumentTransformContext context) {
        return block is DefinitionListBlock definitions
            ? ExpandDefinitionList(definitions, context)
            : new[] { block };
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
                entry.DefinitionBlocks));
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

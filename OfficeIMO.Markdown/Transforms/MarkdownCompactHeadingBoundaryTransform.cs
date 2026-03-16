namespace OfficeIMO.Markdown;

/// <summary>
/// Splits paragraph content when a compact ATX heading marker was emitted directly after prose on the same line.
/// </summary>
/// <remarks>
/// This transform is intended for recoverable paragraph-level transcript/document cleanup where markdown already
/// parsed into a paragraph block, but the paragraph text still contains an inline <c>##</c>-to-<c>######</c> marker
/// that should begin a new heading. The split is performed directly on the parsed inline AST rather than by
/// round-tripping markdown text back through the reader.
/// </remarks>
/// <example>
/// <code>
/// var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
/// options.DocumentTransforms.Add(new MarkdownCompactHeadingBoundaryTransform());
///
/// var document = MarkdownReader.Parse("previous shutdown was unexpected### Reason", options);
/// </code>
/// </example>
public sealed class MarkdownCompactHeadingBoundaryTransform : IMarkdownDocumentTransform {
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
        if (!MarkdownInlineTransformHelpers.TrySplitCompactHeadingBoundary(
                paragraph.Inlines,
                out var leadingParagraph,
                out var headingLevel,
                out var headingTail)) {
            blocks = Array.Empty<IMarkdownBlock>();
            return false;
        }

        leadingParagraph = MarkdownInlineTransformHelpers.TrimWhitespace(leadingParagraph, trimStart: false, trimEnd: true);
        headingTail = MarkdownInlineTransformHelpers.TrimWhitespace(headingTail, trimStart: true, trimEnd: true);
        if (!MarkdownInlineTransformHelpers.HasVisibleContent(leadingParagraph)
            || !MarkdownInlineTransformHelpers.HasVisibleContent(headingTail)) {
            blocks = Array.Empty<IMarkdownBlock>();
            return false;
        }

        var rewritten = new List<IMarkdownBlock> {
            new ParagraphBlock(leadingParagraph)
        };

        AppendHeadingBlocks(rewritten, headingLevel, headingTail);
        blocks = rewritten;
        return rewritten.Count > 1;
    }

    private static void AppendHeadingBlocks(List<IMarkdownBlock> blocks, int level, InlineSequence content) {
        var current = MarkdownInlineTransformHelpers.TrimWhitespace(content, trimStart: true, trimEnd: true);
        if (!MarkdownInlineTransformHelpers.HasVisibleContent(current)) {
            return;
        }

        if (!MarkdownInlineTransformHelpers.TrySplitCompactHeadingBoundary(current, out var head, out var nextLevel, out var tail)) {
            blocks.Add(new HeadingBlock(level, current));
            return;
        }

        head = MarkdownInlineTransformHelpers.TrimWhitespace(head, trimStart: true, trimEnd: true);
        tail = MarkdownInlineTransformHelpers.TrimWhitespace(tail, trimStart: true, trimEnd: true);

        if (!MarkdownInlineTransformHelpers.HasVisibleContent(head)
            || !MarkdownInlineTransformHelpers.HasVisibleContent(tail)) {
            blocks.Add(new HeadingBlock(level, current));
            return;
        }

        blocks.Add(new HeadingBlock(level, head));
        AppendHeadingBlocks(blocks, nextLevel, tail);
    }
}

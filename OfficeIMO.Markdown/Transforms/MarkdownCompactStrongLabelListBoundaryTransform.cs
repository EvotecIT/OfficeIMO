using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Markdown;

/// <summary>
/// Splits compact unordered strong-label list boundaries that were emitted inline after punctuation or symbols.
/// </summary>
/// <remarks>
/// This transform is intended for recoverable cleanup where markdown already parsed into a paragraph block or a
/// simple unordered list item, but the content still contains an inline list marker that should begin a new list item.
/// Code spans are masked during detection so inline code content is preserved.
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
    private static readonly Regex CompactStrongLabelListBoundaryRegex = new(
        @"(?<=[\p{P}\p{S}\)])(?<marker>[-+*])\s+(?=\*\*)",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

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
                case ParagraphBlock paragraph when TryRewriteParagraph(paragraph, context, out var expandedBlocks):
                    rewritten.AddRange(expandedBlocks);
                    break;
                case UnorderedListBlock unordered:
                    RewriteUnorderedList(unordered, context);
                    rewritten.Add(unordered);
                    break;
                default:
                    rewritten.Add(block);
                    break;
            }
        }

        return rewritten;
    }

    private static void RewriteUnorderedList(UnorderedListBlock list, MarkdownDocumentTransformContext context) {
        if (list.Items.Count == 0) {
            return;
        }

        var rewrittenItems = new List<ListItem>(list.Items.Count);
        for (var i = 0; i < list.Items.Count; i++) {
            var item = list.Items[i];
            if (item == null) {
                continue;
            }

            if (TryRewriteSimpleListItem(item, context, out var splitItems)) {
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
        MarkdownDocumentTransformContext context,
        out IReadOnlyList<IMarkdownBlock> blocks) {
        var markdown = paragraph.Inlines.RenderMarkdown();
        if (!TrySplitMarkdown(markdown, paragraph.Inlines, out var headMarkdown, out var marker, out var tailMarkdown)) {
            blocks = Array.Empty<IMarkdownBlock>();
            return false;
        }

        var normalized = headMarkdown + "\n" + marker + " " + tailMarkdown;
        var readerOptions = context.ReaderOptions ?? new MarkdownReaderOptions();
        var rewrittenBlocks = MarkdownReader.ParseBlockFragment(normalized, readerOptions, new MarkdownReaderState());
        if (rewrittenBlocks.Count == 1
            && rewrittenBlocks[0] is ParagraphBlock rewrittenParagraph
            && rewrittenParagraph.Inlines.RenderMarkdown().Equals(markdown, StringComparison.Ordinal)) {
            blocks = Array.Empty<IMarkdownBlock>();
            return false;
        }

        blocks = rewrittenBlocks;
        return rewrittenBlocks.Count > 0;
    }

    private static bool TryRewriteSimpleListItem(
        ListItem item,
        MarkdownDocumentTransformContext context,
        out IReadOnlyList<ListItem> items) {
        if (item.IsTask || item.AdditionalParagraphs.Count > 0 || item.Children.Count > 0) {
            items = Array.Empty<ListItem>();
            return false;
        }

        var markdown = item.Content.RenderMarkdown();
        if (!TrySplitMarkdown(markdown, item.Content, out var headMarkdown, out var marker, out var tailMarkdown)) {
            items = Array.Empty<ListItem>();
            return false;
        }

        var readerOptions = context.ReaderOptions ?? new MarkdownReaderOptions();
        var parsedTailBlocks = MarkdownReader.ParseBlockFragment(marker + " " + tailMarkdown, readerOptions, new MarkdownReaderState());
        if (parsedTailBlocks.Count != 1 || parsedTailBlocks[0] is not UnorderedListBlock tailList || tailList.Items.Count == 0) {
            items = Array.Empty<ListItem>();
            return false;
        }

        var rewritten = new List<ListItem>(tailList.Items.Count + 1) {
            new ListItem(MarkdownReader.ParseInlineText(headMarkdown.TrimEnd(), readerOptions, new MarkdownReaderState())) {
                Level = item.Level,
                ForceLoose = item.ForceLoose
            }
        };

        for (var i = 0; i < tailList.Items.Count; i++) {
            tailList.Items[i].Level = item.Level;
            rewritten.Add(tailList.Items[i]);
        }

        items = rewritten;
        return true;
    }

    private static bool TrySplitMarkdown(
        string markdown,
        InlineSequence inlines,
        out string headMarkdown,
        out string marker,
        out string tailMarkdown) {
        headMarkdown = string.Empty;
        marker = string.Empty;
        tailMarkdown = string.Empty;

        if (markdown.Length == 0) {
            return false;
        }

        var maskedMarkdown = RenderMaskedMarkdown(inlines);
        var match = CompactStrongLabelListBoundaryRegex.Match(maskedMarkdown);
        if (!match.Success) {
            return false;
        }

        headMarkdown = markdown.Substring(0, match.Index).TrimEnd();
        marker = match.Groups["marker"].Value;
        tailMarkdown = markdown.Substring(match.Index + match.Length);
        return headMarkdown.Length > 0 && marker.Length == 1 && tailMarkdown.Length > 0;
    }

    private static string RenderMaskedMarkdown(InlineSequence inlines) {
        var sb = new StringBuilder();
        var nodes = inlines?.Nodes;
        if (nodes == null || nodes.Count == 0) {
            return string.Empty;
        }

        for (var i = 0; i < nodes.Count; i++) {
            var node = nodes[i];
            if (node == null) {
                continue;
            }

            var rendered = ((IRenderableMarkdownInline)node).RenderMarkdown();
            if (rendered.Length == 0) {
                continue;
            }

            if (node is CodeSpanInline) {
                sb.Append(' ', rendered.Length);
                continue;
            }

            sb.Append(rendered);
        }

        return sb.ToString();
    }
}

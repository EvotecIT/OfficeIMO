using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Markdown;

/// <summary>
/// Splits paragraph content when a prose label ending with a colon was emitted directly before a list marker.
/// </summary>
/// <remarks>
/// This transform is intended for recoverable paragraph-level cleanup where markdown already parsed into a
/// paragraph block, but the paragraph text still contains an inline list marker that should start a new list item.
/// Code spans are masked during detection so inline code content is preserved.
/// </remarks>
/// <example>
/// <code>
/// var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
/// options.DocumentTransforms.Add(new MarkdownColonListBoundaryTransform());
///
/// var document = MarkdownReader.Parse("Next step:- **Item**", options);
/// </code>
/// </example>
public sealed class MarkdownColonListBoundaryTransform : IMarkdownDocumentTransform {
    private static readonly Regex ColonListBoundaryRegex = new(
        @":\s*(?<marker>[-+*])\s+(?=(\*\*|`|\[|\p{L}|\p{N}))",
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

            if (block is ParagraphBlock paragraph
                && TryRewriteParagraph(paragraph, context, out var expandedBlocks)) {
                rewritten.AddRange(expandedBlocks);
                continue;
            }

            rewritten.Add(block);
        }

        return rewritten;
    }

    private static bool TryRewriteParagraph(
        ParagraphBlock paragraph,
        MarkdownDocumentTransformContext context,
        out IReadOnlyList<IMarkdownBlock> blocks) {
        var markdown = paragraph.Inlines.RenderMarkdown();
        if (markdown.Length == 0 || markdown.IndexOf(':') < 0) {
            blocks = Array.Empty<IMarkdownBlock>();
            return false;
        }

        var maskedMarkdown = RenderMaskedMarkdown(paragraph.Inlines);
        var match = ColonListBoundaryRegex.Match(maskedMarkdown);
        if (!match.Success) {
            blocks = Array.Empty<IMarkdownBlock>();
            return false;
        }

        var normalized = markdown.Substring(0, match.Index + 1)
                        + "\n"
                        + match.Groups["marker"].Value
                        + " "
                        + markdown.Substring(match.Index + match.Length);

        if (normalized.Equals(markdown, StringComparison.Ordinal)) {
            blocks = Array.Empty<IMarkdownBlock>();
            return false;
        }

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

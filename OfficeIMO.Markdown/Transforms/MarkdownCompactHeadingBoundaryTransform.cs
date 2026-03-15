using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Markdown;

/// <summary>
/// Splits paragraph content when a compact ATX heading marker was emitted directly after prose on the same line.
/// </summary>
/// <remarks>
/// This transform is intended for recoverable paragraph-level transcript/document cleanup where markdown already
/// parsed into a paragraph block, but the paragraph text still contains an inline <c>##</c>-to-<c>######</c> marker
/// that should begin a new heading. Code spans are masked during detection so inline code content is preserved.
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
    private static readonly Regex CompactHeadingBoundaryRegex = new(
        @"(?<=[^\s\r\n])(?<marker>#{2,6})\s+(?=\S)",
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
        if (markdown.Length == 0 || markdown.IndexOf('#') < 0) {
            blocks = Array.Empty<IMarkdownBlock>();
            return false;
        }

        var maskedMarkdown = RenderMaskedMarkdown(paragraph.Inlines);
        var matches = CompactHeadingBoundaryRegex.Matches(maskedMarkdown);
        if (matches.Count == 0) {
            blocks = Array.Empty<IMarkdownBlock>();
            return false;
        }

        var normalized = ApplyBoundaryRewrites(markdown, matches);
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

    private static string ApplyBoundaryRewrites(string markdown, MatchCollection matches) {
        var rewritten = markdown;
        for (var i = matches.Count - 1; i >= 0; i--) {
            var match = matches[i];
            if (!match.Success) {
                continue;
            }

            var marker = match.Groups["marker"].Value;
            if (marker.Length < 2 || marker.Length > 6) {
                continue;
            }

            rewritten = rewritten.Substring(0, match.Index)
                        + "\n"
                        + marker
                        + " "
                        + rewritten.Substring(match.Index + match.Length);
        }

        return rewritten;
    }
}

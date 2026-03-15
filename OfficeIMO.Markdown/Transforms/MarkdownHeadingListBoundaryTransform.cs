using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Markdown;

/// <summary>
/// Splits heading content when an unordered strong-label list marker was emitted directly after the heading text.
/// </summary>
/// <remarks>
/// This transform is intended for recoverable heading-level cleanup where markdown already parsed into a heading
/// block, but the heading text still contains an inline list marker that should begin a new unordered list item.
/// Code spans are masked during detection so inline code content is preserved.
/// </remarks>
/// <example>
/// <code>
/// var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
/// options.DocumentTransforms.Add(new MarkdownHeadingListBoundaryTransform());
///
/// var document = MarkdownReader.Parse("## Summary- **Item:** value", options);
/// </code>
/// </example>
public sealed class MarkdownHeadingListBoundaryTransform : IMarkdownDocumentTransform {
    private static readonly Regex HeadingListBoundaryRegex = new(
        @"(?<!\s)(?<marker>[-+*])\s+(?=\*\*)",
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

            if (block is HeadingBlock heading
                && TryRewriteHeading(heading, context, out var expandedBlocks)) {
                rewritten.AddRange(expandedBlocks);
                continue;
            }

            rewritten.Add(block);
        }

        return rewritten;
    }

    private static bool TryRewriteHeading(
        HeadingBlock heading,
        MarkdownDocumentTransformContext context,
        out IReadOnlyList<IMarkdownBlock> blocks) {
        var headingMarkdown = heading.Inlines.RenderMarkdown();
        if (headingMarkdown.Length == 0 || headingMarkdown.IndexOf('-') < 0 && headingMarkdown.IndexOf('+') < 0 && headingMarkdown.IndexOf('*') < 0) {
            blocks = Array.Empty<IMarkdownBlock>();
            return false;
        }

        var maskedMarkdown = RenderMaskedMarkdown(heading.Inlines);
        var match = HeadingListBoundaryRegex.Match(maskedMarkdown);
        if (!match.Success) {
            blocks = Array.Empty<IMarkdownBlock>();
            return false;
        }

        var headingText = headingMarkdown.Substring(0, match.Index).TrimEnd();
        if (headingText.Length == 0) {
            blocks = Array.Empty<IMarkdownBlock>();
            return false;
        }

        var listMarkdown = match.Groups["marker"].Value + " " + headingMarkdown.Substring(match.Index + match.Length);
        var readerOptions = context.ReaderOptions ?? new MarkdownReaderOptions();
        var parsedListBlocks = MarkdownReader.ParseBlockFragment(listMarkdown, readerOptions, new MarkdownReaderState());
        if (parsedListBlocks.Count == 0) {
            blocks = Array.Empty<IMarkdownBlock>();
            return false;
        }

        var rewritten = new List<IMarkdownBlock>(parsedListBlocks.Count + 1) {
            new HeadingBlock(heading.Level, MarkdownReader.ParseInlineText(headingText, readerOptions, new MarkdownReaderState()))
        };

        for (var i = 0; i < parsedListBlocks.Count; i++) {
            rewritten.Add(parsedListBlocks[i]);
        }

        blocks = rewritten;
        return true;
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

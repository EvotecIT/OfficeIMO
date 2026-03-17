using System.Text.RegularExpressions;

namespace OfficeIMO.Markdown;

/// <summary>
/// Removes cached-evidence transport marker paragraphs emitted in IntelligenceX transcript flows.
/// </summary>
/// <remarks>
/// This intentionally runs as a document transform because the marker is already parseable markdown and
/// should be removed structurally rather than through pre-parse text surgery.
/// </remarks>
public sealed class MarkdownIntelligenceXCachedToolEvidenceMarkerTransform : IMarkdownDocumentTransform {
    private static readonly Regex CachedToolEvidenceMarkerParagraphRegex = new(
        @"^ix\s*:\s*cached-tool-evidence\s*:\s*v1$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

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

            if (block is ParagraphBlock paragraph && IsCachedToolEvidenceMarker(paragraph)) {
                continue;
            }

            rewritten.Add(block);
        }

        return rewritten;
    }

    private static bool IsCachedToolEvidenceMarker(ParagraphBlock paragraph) {
        var markdown = ((IMarkdownBlock)paragraph).RenderMarkdown().Trim();
        return CachedToolEvidenceMarkerParagraphRegex.IsMatch(markdown);
    }
}

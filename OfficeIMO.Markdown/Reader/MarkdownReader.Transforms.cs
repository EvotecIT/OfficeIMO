using System.IO;
using System.Linq;
using System.Text;
// Intentionally avoid heavy regex use; simple scanning is used for resilience and speed.

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static MarkdownDoc ApplyDocumentTransforms(
        MarkdownDoc document,
        MarkdownReaderOptions options,
        ICollection<MarkdownDocumentTransformDiagnostic>? diagnostics = null,
        MarkdownSyntaxNode? syntaxTree = null,
        string? sourceMarkdown = null,
        string? originalMarkdown = null,
        bool preservesOriginalMarkdown = false) {
        var transforms = BuildEffectiveDocumentTransforms(options);
        var topLevelBlockSourceSpans = syntaxTree == null
            ? null
            : BuildTopLevelBlockSourceSpans(document, syntaxTree);
        return MarkdownDocumentTransformPipeline.Apply(
            document,
            transforms,
            new MarkdownDocumentTransformContext(
                MarkdownDocumentTransformSource.MarkdownReader,
                options,
                sourceOptions: null,
                diagnostics,
                syntaxTree,
                topLevelBlockSourceSpans,
                sourceMarkdown,
                originalMarkdown,
                preservesOriginalMarkdown));
    }

    private static IReadOnlyList<MarkdownSourceSpan?> BuildTopLevelBlockSourceSpans(MarkdownDoc document, MarkdownSyntaxNode syntaxTree) {
        var spans = new List<MarkdownSourceSpan?>(document.Blocks.Count);
        var blockChildren = syntaxTree.Children
            .Where(static child => child.AssociatedObject is IMarkdownBlock)
            .ToList();
        var topLevelBlocks = document.TopLevelBlocks;
        var childCount = Math.Min(blockChildren.Count, topLevelBlocks.Count);
        for (var i = 0; i < childCount; i++) {
            if (topLevelBlocks[i] is FrontMatterBlock) {
                continue;
            }

            spans.Add(blockChildren[i].SourceSpan);
        }

        while (spans.Count < document.Blocks.Count) {
            spans.Add(null);
        }

        if (spans.Count > document.Blocks.Count) {
            spans.RemoveRange(document.Blocks.Count, spans.Count - document.Blocks.Count);
        }

        return spans;
    }

    private static IReadOnlyList<IMarkdownDocumentTransform> BuildEffectiveDocumentTransforms(MarkdownReaderOptions options) {
        if (options == null) {
            return Array.Empty<IMarkdownDocumentTransform>();
        }

        var normalization = options.InputNormalization;
        bool needsRegisteredFencedBlockTransform = options.FencedBlockExtensions.Count > 0;
        // These flags intentionally map to AST/document transforms rather than pre-parse text
        // repair because the markdown already parses into recoverable paragraph/heading/list
        // structures. Keeping them here makes the routing boundary explicit and prevents the
        // pre-parse normalizer from growing back into a general transcript rewrite pipeline.
        bool needsStandaloneHashTransform = normalization?.NormalizeStandaloneHashHeadingSeparators == true;
        bool needsCompactHeadingBoundaryTransform = normalization?.NormalizeCompactHeadingBoundaries == true;
        bool needsColonListBoundaryTransform = normalization?.NormalizeColonListBoundaries == true;
        bool needsHeadingListBoundaryTransform = normalization?.NormalizeHeadingListBoundaries == true;
        bool needsCompactStrongLabelListBoundaryTransform = normalization?.NormalizeCompactStrongLabelListBoundaries == true;
        bool needsListStrongArtifactTransform =
            normalization?.NormalizeLooseStrongDelimiters == true
            || normalization?.NormalizeDanglingTrailingStrongListClosers == true
            || normalization?.NormalizeMetricValueStrongRuns == true;

        if (!needsRegisteredFencedBlockTransform
            && !needsStandaloneHashTransform
            && !needsCompactHeadingBoundaryTransform
            && !needsColonListBoundaryTransform
            && !needsHeadingListBoundaryTransform
            && !needsCompactStrongLabelListBoundaryTransform
            && !needsListStrongArtifactTransform) {
            return options.DocumentTransforms;
        }

        var configured = options.DocumentTransforms;
        bool hasStandaloneHashTransform = false;
        bool hasCompactHeadingBoundaryTransform = false;
        bool hasColonListBoundaryTransform = false;
        bool hasHeadingListBoundaryTransform = false;
        bool hasCompactStrongLabelListBoundaryTransform = false;
        bool hasListStrongArtifactTransform = false;
        bool hasRegisteredFencedBlockTransform = false;

        for (var i = 0; i < configured.Count; i++) {
            switch (configured[i]) {
                case MarkdownRegisteredFencedBlockTransform:
                    hasRegisteredFencedBlockTransform = true;
                    break;
                case MarkdownStandaloneHashHeadingSeparatorTransform:
                    hasStandaloneHashTransform = true;
                    break;
                case MarkdownCompactHeadingBoundaryTransform:
                    hasCompactHeadingBoundaryTransform = true;
                    break;
                case MarkdownColonListBoundaryTransform:
                    hasColonListBoundaryTransform = true;
                    break;
                case MarkdownHeadingListBoundaryTransform:
                    hasHeadingListBoundaryTransform = true;
                    break;
                case MarkdownCompactStrongLabelListBoundaryTransform:
                    hasCompactStrongLabelListBoundaryTransform = true;
                    break;
                case MarkdownListParagraphStrongArtifactTransform:
                    hasListStrongArtifactTransform = true;
                    break;
            }
        }

        if ((!needsRegisteredFencedBlockTransform || hasRegisteredFencedBlockTransform)
            && (!needsStandaloneHashTransform || hasStandaloneHashTransform)
            && (!needsCompactHeadingBoundaryTransform || hasCompactHeadingBoundaryTransform)
            && (!needsColonListBoundaryTransform || hasColonListBoundaryTransform)
            && (!needsHeadingListBoundaryTransform || hasHeadingListBoundaryTransform)
            && (!needsCompactStrongLabelListBoundaryTransform || hasCompactStrongLabelListBoundaryTransform)
            && (!needsListStrongArtifactTransform || hasListStrongArtifactTransform)) {
            return configured;
        }

        var transforms = new List<IMarkdownDocumentTransform>(configured.Count + 7);
        if (needsRegisteredFencedBlockTransform && !hasRegisteredFencedBlockTransform) {
            transforms.Add(new MarkdownRegisteredFencedBlockTransform(options.FencedBlockExtensions));
        }

        if (needsListStrongArtifactTransform && !hasListStrongArtifactTransform) {
            transforms.Add(new MarkdownListParagraphStrongArtifactTransform(normalization!));
        }

        if (needsCompactHeadingBoundaryTransform && !hasCompactHeadingBoundaryTransform) {
            transforms.Add(new MarkdownCompactHeadingBoundaryTransform());
        }

        if (needsHeadingListBoundaryTransform && !hasHeadingListBoundaryTransform) {
            transforms.Add(new MarkdownHeadingListBoundaryTransform());
        }

        if (needsCompactStrongLabelListBoundaryTransform && !hasCompactStrongLabelListBoundaryTransform) {
            transforms.Add(new MarkdownCompactStrongLabelListBoundaryTransform());
        }

        if (needsColonListBoundaryTransform && !hasColonListBoundaryTransform) {
            transforms.Add(new MarkdownColonListBoundaryTransform());
        }

        if (needsStandaloneHashTransform && !hasStandaloneHashTransform) {
            transforms.Add(new MarkdownStandaloneHashHeadingSeparatorTransform());
        }

        for (var i = 0; i < configured.Count; i++) {
            if (configured[i] != null) {
                transforms.Add(configured[i]);
            }
        }

        return transforms;
    }
}

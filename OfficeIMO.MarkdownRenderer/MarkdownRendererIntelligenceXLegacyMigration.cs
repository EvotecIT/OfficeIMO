using System.Collections.Generic;
using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Legacy markdown migration helpers for older IntelligenceX transcript artifacts.
/// This owns compatibility cleanup only; alias fence registration stays in <see cref="MarkdownRendererIntelligenceXAdapter"/>.
/// </summary>
public static class MarkdownRendererIntelligenceXLegacyMigration {
    private static readonly MarkdownIntelligenceXCachedToolEvidenceMarkerTransform CachedToolEvidenceMarkerTransform = new();
    private static readonly MarkdownIntelligenceXLegacyToolHeadingTransform LegacyToolHeadingTransform = new();

    /// <summary>
    /// Registers the legacy IX transcript compatibility helpers if they are not already present.
    /// </summary>
    public static void Apply(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        AddReaderTransformIfMissing(options, CachedToolEvidenceMarkerTransform);
        AddReaderTransformIfMissing(options, LegacyToolHeadingTransform);
    }

    /// <summary>
    /// Returns <see langword="true"/> when any IX legacy migration helper is present.
    /// </summary>
    public static bool IsApplied(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        var transforms = options.ReaderOptions.DocumentTransforms;
        for (var i = 0; i < transforms.Count; i++) {
            if (transforms[i] is MarkdownIntelligenceXCachedToolEvidenceMarkerTransform
                || transforms[i] is MarkdownIntelligenceXLegacyToolHeadingTransform) {
                return true;
            }
        }

        return false;
    }

    private static void AddReaderTransformIfMissing(MarkdownRendererOptions options, IMarkdownDocumentTransform transform) {
        var transforms = options.ReaderOptions.DocumentTransforms;
        for (var i = 0; i < transforms.Count; i++) {
            if ((transform is MarkdownIntelligenceXCachedToolEvidenceMarkerTransform
                    && transforms[i] is MarkdownIntelligenceXCachedToolEvidenceMarkerTransform)
                || (transform is MarkdownIntelligenceXLegacyToolHeadingTransform
                    && transforms[i] is MarkdownIntelligenceXLegacyToolHeadingTransform)) {
                return;
            }
        }

        transforms.Add(transform);
    }
}

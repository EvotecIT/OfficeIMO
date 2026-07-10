using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Reader;

/// <summary>Filters shared assets and removes OCR candidates that target removed asset ids.</summary>
public sealed class OfficeDocumentAssetFilterProcessor : OfficeDocumentProcessorBase {
    private readonly Func<OfficeDocumentAsset, bool> _predicate;

    /// <summary>
    /// Creates the processor. The predicate must be deterministic and safe for concurrent calls when used by a shared reader.
    /// </summary>
    public OfficeDocumentAssetFilterProcessor(Func<OfficeDocumentAsset, bool> predicate)
        : base("officeimo.reader.filter-assets") {
        _predicate = predicate ?? throw new ArgumentNullException(nameof(predicate));
    }

    /// <inheritdoc />
    public override OfficeDocumentReadResult Process(
        OfficeDocumentReadResult document,
        OfficeDocumentProcessorContext context) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        var removedIds = new HashSet<string>(StringComparer.Ordinal);
        document.Assets = FilterAssets(document.Assets, removedIds, context.CancellationToken);
        foreach (OfficeDocumentPage page in document.Pages ?? Array.Empty<OfficeDocumentPage>()) {
            context.CancellationToken.ThrowIfCancellationRequested();
            page.Assets = FilterAssets(page.Assets, removedIds, context.CancellationToken);
        }

        if (removedIds.Count == 0) return document;
        document.OcrCandidates = FilterCandidates(document.OcrCandidates, removedIds);
        foreach (OfficeDocumentPage page in document.Pages ?? Array.Empty<OfficeDocumentPage>()) {
            page.OcrCandidates = FilterCandidates(page.OcrCandidates, removedIds);
        }
        return document;
    }

    private IReadOnlyList<OfficeDocumentAsset> FilterAssets(
        IReadOnlyList<OfficeDocumentAsset>? assets,
        ISet<string> removedIds,
        System.Threading.CancellationToken cancellationToken) {
        if (assets == null || assets.Count == 0) return Array.Empty<OfficeDocumentAsset>();
        var kept = new List<OfficeDocumentAsset>(assets.Count);
        for (int index = 0; index < assets.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeDocumentAsset asset = assets[index];
            if (asset != null && _predicate(asset)) {
                kept.Add(asset);
            } else if (asset != null && !string.IsNullOrWhiteSpace(asset.Id)) {
                removedIds.Add(asset.Id);
            }
        }
        return kept.Count == 0 ? Array.Empty<OfficeDocumentAsset>() : kept.ToArray();
    }

    private static IReadOnlyList<OfficeDocumentOcrCandidate> FilterCandidates(
        IReadOnlyList<OfficeDocumentOcrCandidate>? candidates,
        ISet<string> removedIds) {
        if (candidates == null || candidates.Count == 0) return Array.Empty<OfficeDocumentOcrCandidate>();
        return candidates
            .Where(candidate => candidate != null &&
                (string.IsNullOrWhiteSpace(candidate.AssetId) || !removedIds.Contains(candidate.AssetId!)))
            .ToArray();
    }
}

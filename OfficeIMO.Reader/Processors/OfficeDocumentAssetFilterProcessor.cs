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
        var keptIds = new HashSet<string>(StringComparer.Ordinal);
        document.Assets = FilterAssets(document.Assets, removedIds, keptIds, context.CancellationToken);
        foreach (OfficeDocumentPage page in document.Pages ?? Array.Empty<OfficeDocumentPage>()) {
            context.CancellationToken.ThrowIfCancellationRequested();
            page.Assets = FilterAssets(page.Assets, removedIds, keptIds, context.CancellationToken);
        }
        document.Metadata = RefreshAssetCountMetadata(document.Metadata, document.Assets.Count);

        removedIds.ExceptWith(keptIds);
        if (removedIds.Count == 0) return document;
        IReadOnlyList<OfficeDocumentOcrCandidate> removedCandidates = EnumerateCandidates(document)
            .Where(candidate => !string.IsNullOrWhiteSpace(candidate.AssetId) && removedIds.Contains(candidate.AssetId!))
            .ToArray();
        document.OcrCandidates = FilterCandidates(document.OcrCandidates, removedIds);
        foreach (OfficeDocumentPage page in document.Pages ?? Array.Empty<OfficeDocumentPage>()) {
            page.OcrCandidates = FilterCandidates(page.OcrCandidates, removedIds);
        }
        document.Diagnostics = FilterOcrDiagnostics(document, removedIds, removedCandidates);
        return document;
    }

    private static IReadOnlyList<OfficeDocumentMetadataEntry> RefreshAssetCountMetadata(
        IReadOnlyList<OfficeDocumentMetadataEntry>? metadata,
        int assetCount) {
        if (metadata == null || metadata.Count == 0) return Array.Empty<OfficeDocumentMetadataEntry>();
        var refreshed = new List<OfficeDocumentMetadataEntry>(metadata.Count);
        bool emittedCount = false;
        foreach (OfficeDocumentMetadataEntry entry in metadata) {
            if (entry == null) continue;
            if (!string.Equals(entry.Id, "reader-asset-count", StringComparison.Ordinal)) {
                refreshed.Add(entry);
                continue;
            }
            if (emittedCount || assetCount == 0) continue;
            entry.Value = assetCount.ToString(System.Globalization.CultureInfo.InvariantCulture);
            refreshed.Add(entry);
            emittedCount = true;
        }
        return refreshed.Count == 0 ? Array.Empty<OfficeDocumentMetadataEntry>() : refreshed.ToArray();
    }

    private IReadOnlyList<OfficeDocumentAsset> FilterAssets(
        IReadOnlyList<OfficeDocumentAsset>? assets,
        ISet<string> removedIds,
        ISet<string> keptIds,
        System.Threading.CancellationToken cancellationToken) {
        if (assets == null || assets.Count == 0) return Array.Empty<OfficeDocumentAsset>();
        var kept = new List<OfficeDocumentAsset>(assets.Count);
        for (int index = 0; index < assets.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeDocumentAsset asset = assets[index];
            if (asset != null && _predicate(asset)) {
                kept.Add(asset);
                if (!string.IsNullOrWhiteSpace(asset.Id)) keptIds.Add(asset.Id);
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

    private static IReadOnlyList<OfficeDocumentDiagnostic> FilterOcrDiagnostics(
        OfficeDocumentReadResult document,
        ISet<string> removedIds,
        IReadOnlyList<OfficeDocumentOcrCandidate> removedCandidates) {
        var removedSignatures = new HashSet<string>(
            removedCandidates.Select(candidate => BuildOcrSignature(candidate.Reason, candidate.Location)),
            StringComparer.Ordinal);
        var remainingSignatures = new HashSet<string>(
            EnumerateCandidates(document).Select(candidate => BuildOcrSignature(candidate.Reason, candidate.Location)),
            StringComparer.Ordinal);
        var diagnostics = new List<OfficeDocumentDiagnostic>();
        foreach (OfficeDocumentDiagnostic diagnostic in document.Diagnostics ?? Array.Empty<OfficeDocumentDiagnostic>()) {
            if (diagnostic == null) continue;
            if (!string.Equals(diagnostic.Code, "ocr-needed", StringComparison.Ordinal)) {
                diagnostics.Add(diagnostic);
                continue;
            }
            if (IsTiedToRemovedAsset(diagnostic, removedIds)) continue;
            string signature = BuildOcrSignature(diagnostic.Message, diagnostic.Location);
            if (removedSignatures.Contains(signature) && !remainingSignatures.Contains(signature)) continue;
            diagnostics.Add(diagnostic);
        }
        return diagnostics.Count == 0 ? Array.Empty<OfficeDocumentDiagnostic>() : diagnostics.ToArray();
    }

    private static bool IsTiedToRemovedAsset(OfficeDocumentDiagnostic diagnostic, ISet<string> removedIds) {
        if (!string.IsNullOrWhiteSpace(diagnostic.Location?.BlockAnchor) &&
            removedIds.Contains(diagnostic.Location!.BlockAnchor!)) return true;
        return diagnostic.Attributes != null &&
            diagnostic.Attributes.TryGetValue("assetId", out string? assetId) &&
            !string.IsNullOrWhiteSpace(assetId) &&
            removedIds.Contains(assetId!);
    }

    private static string BuildOcrSignature(string? reason, ReaderLocation? location) => string.Join("|", new[] {
        reason ?? string.Empty,
        location?.Path ?? string.Empty,
        location?.Page?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty,
        location?.Slide?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty,
        location?.Sheet ?? string.Empty,
        location?.BlockAnchor ?? string.Empty
    });

    private static IEnumerable<OfficeDocumentOcrCandidate> EnumerateCandidates(OfficeDocumentReadResult document) {
        foreach (OfficeDocumentOcrCandidate candidate in document.OcrCandidates ?? Array.Empty<OfficeDocumentOcrCandidate>()) {
            if (candidate != null) yield return candidate;
        }
        foreach (OfficeDocumentPage page in document.Pages ?? Array.Empty<OfficeDocumentPage>()) {
            if (page?.OcrCandidates == null) continue;
            foreach (OfficeDocumentOcrCandidate candidate in page.OcrCandidates) {
                if (candidate != null) yield return candidate;
            }
        }
    }
}

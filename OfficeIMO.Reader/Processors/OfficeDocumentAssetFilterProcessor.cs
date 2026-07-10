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

        removedIds.ExceptWith(keptIds);
        if (removedIds.Count == 0) return document;
        document.OcrCandidates = FilterCandidates(document.OcrCandidates, removedIds);
        foreach (OfficeDocumentPage page in document.Pages ?? Array.Empty<OfficeDocumentPage>()) {
            page.OcrCandidates = FilterCandidates(page.OcrCandidates, removedIds);
        }
        document.Diagnostics = RebuildOcrDiagnostics(document);
        return document;
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

    private static IReadOnlyList<OfficeDocumentDiagnostic> RebuildOcrDiagnostics(OfficeDocumentReadResult document) {
        var diagnostics = (document.Diagnostics ?? Array.Empty<OfficeDocumentDiagnostic>())
            .Where(static diagnostic => diagnostic != null &&
                !string.Equals(diagnostic.Code, "ocr-needed", StringComparison.Ordinal))
            .ToList();
        var seenIds = new HashSet<string>(StringComparer.Ordinal);
        var seenCandidates = new HashSet<OfficeDocumentOcrCandidate>(ReferenceIdentityComparer<OfficeDocumentOcrCandidate>.Instance);
        foreach (OfficeDocumentOcrCandidate candidate in EnumerateCandidates(document)) {
            if (candidate == null || !seenCandidates.Add(candidate)) continue;
            if (!string.IsNullOrWhiteSpace(candidate.Id) && !seenIds.Add(candidate.Id)) continue;
            diagnostics.Add(new OfficeDocumentDiagnostic {
                Severity = OfficeDocumentDiagnosticSeverity.Warning,
                Category = OfficeDocumentDiagnosticCategory.Ocr,
                Code = "ocr-needed",
                Message = candidate.Reason ?? "OCR may be needed before text extraction is complete.",
                Source = "officeimo.reader.filter-assets",
                IsRecoverable = true,
                Location = candidate.Location
            });
        }
        return diagnostics.Count == 0 ? Array.Empty<OfficeDocumentDiagnostic>() : diagnostics.ToArray();
    }

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

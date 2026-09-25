using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Coordinates page alignment, retained visual proof, and supported semantic changes.</summary>
internal static class PdfReviewComparer {
    internal static PdfReviewComparisonReport Compare(
        PdfDocument expectedSource,
        PdfDocument actualSource,
        PdfReviewComparisonOptions? options,
        CancellationToken cancellationToken) {
        Guard.NotNull(expectedSource, nameof(expectedSource));
        Guard.NotNull(actualSource, nameof(actualSource));
        PdfReviewComparisonOptions effective = options ?? new PdfReviewComparisonOptions();
        effective.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        PdfReadDocument expected = expectedSource.GetReadSnapshot(cancellationToken: cancellationToken).Document;
        PdfReadDocument actual = actualSource.GetReadSnapshot(cancellationToken: cancellationToken).Document;
        PdfPageChangeReport alignment = PdfPageChangeAnalyzer.Analyze(expected, actual, effective.PageAlignment,
            effective.Visual.IgnoredRegions.ToArray(), cancellationToken);
        PdfPageChange[] pairs = alignment.Changes.Where(static change => change.ExpectedPageNumber.HasValue && change.ActualPageNumber.HasValue).ToArray();
        if (pairs.Length > effective.MaxAlignedPagePairs) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts, effective.MaxAlignedPagePairs, pairs.Length);
        }
        int changedPairCount = pairs.Count(static pair => pair.Kind == PdfPageChangeKind.ModifiedCandidate);
        int maximumVisualPairs = Math.Min(effective.MaxChangedPagePairs, effective.Visual.MaxPages);
        if (changedPairCount > maximumVisualPairs) throw PdfReadLimitException.Create(PdfReadLimitKind.RenderPages, maximumVisualPairs, changedPairCount);
        if (pairs.Length == 0) return new PdfReviewComparisonReport(alignment, Array.Empty<PdfReviewPageComparison>());

        PdfDocumentReadResult expectedLogical = PdfDocumentReadEngine.Read(expected, new PdfReadOptions {
            Profile = PdfReadProfile.Fast,
            PageSelection = PdfPageSelection.From(pairs.Select(static pair => pair.ExpectedPageNumber!.Value).ToArray())
        }, cancellationToken);
        PdfDocumentReadResult actualLogical = PdfDocumentReadEngine.Read(actual, new PdfReadOptions {
            Profile = PdfReadProfile.Fast,
            PageSelection = PdfPageSelection.From(pairs.Select(static pair => pair.ActualPageNumber!.Value).ToArray())
        }, cancellationToken);

        long totalPixels = 0;
        long totalOutputBytes = 0;
        var pages = new List<PdfReviewPageComparison>(pairs.Length);
        foreach (PdfPageChange pair in pairs) {
            cancellationToken.ThrowIfCancellationRequested();
            int expectedNumber = pair.ExpectedPageNumber!.Value;
            int actualNumber = pair.ActualPageNumber!.Value;
            PdfVisualPageComparison? visual = pair.Kind == PdfPageChangeKind.ModifiedCandidate
                ? PdfVisualComparer.ComparePages(expected, expectedNumber, actual, actualNumber, effective.Visual, ref totalPixels, cancellationToken)
                : null;
            if (visual is not null) {
                totalOutputBytes = checked(totalOutputBytes + visual.OutputByteLength);
                if (totalOutputBytes > effective.Visual.MaxTotalOutputBytes) {
                    throw PdfReadLimitException.Create(PdfReadLimitKind.RenderBytes, effective.Visual.MaxTotalOutputBytes, totalOutputBytes);
                }
            }
            PdfLogicalPage expectedPage = expectedLogical.PagesBySourcePageNumber[expectedNumber][0];
            PdfLogicalPage actualPage = actualLogical.PagesBySourcePageNumber[actualNumber][0];
            IReadOnlyList<PdfReviewChange> changes = PdfReviewSemanticComparer.Compare(
                expectedPage, actualPage, visual, effective, cancellationToken);
            var page = new PdfReviewPageComparison(pair, visual, changes);
            if (!page.IsMatch) {
                if (pair.Kind != PdfPageChangeKind.ModifiedCandidate && ++changedPairCount > effective.MaxChangedPagePairs) {
                    throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts, effective.MaxChangedPagePairs, changedPairCount);
                }
                pages.Add(page);
            }
        }
        return new PdfReviewComparisonReport(alignment, pages);
    }
}

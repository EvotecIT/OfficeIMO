using OfficeIMO.Ocr;
using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed record PdfWorkspaceOrientationReview(PdfWorkspace Owner, long Revision, PdfPageOrientationReview Review);

internal sealed partial class PdfWorkspace {
    internal async Task<PdfWorkspaceOrientationReview> PrepareOrientationReviewAsync(IOcrEngine engine,
        IReadOnlyList<int> pages, CancellationToken token) {
        byte[] bytes; long revision;
        await _operationGate.WaitAsync(token).ConfigureAwait(false);
        try { ThrowIfDisposed(); token.ThrowIfCancellationRequested(); bytes = _bytes; revision = Revision; }
        finally { _operationGate.Release(); }
        var options = new PdfOcrMergeOptions { ReadOptions = new() { PageSelection = OfficeIMO.Pdf.PdfPageSelection.From(pages.ToArray()) } };
        var review = await RunNonDetachableCpuWorkAsync(
            () => LoadDocument(bytes).PrepareOrientationReviewAsync(engine, options, token), token).ConfigureAwait(false);
        token.ThrowIfCancellationRequested(); ThrowIfDisposed();
        if (Revision != revision) throw new InvalidOperationException("The document changed during orientation detection.");
        return new(this, revision, review);
    }

    internal Task ApplyOrientationReviewAsync(PdfWorkspaceOrientationReview review,
        IReadOnlyList<PdfPageOrientationProposal> accepted, CancellationToken token,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        if (!ReferenceEquals(review.Owner, this)) throw new ArgumentException("The review belongs to another document.", nameof(review));
        var captured = accepted.ToArray();
        if (captured.Length == 0) throw new ArgumentException("Accept at least one correction.", nameof(accepted));
        return MutateBytesAsync(PdfWorkspaceOperationKind.Rotate, "Applied reviewed page orientation corrections",
            captured.Select(page => page.PageNumber).ToArray(), bytes => {
                if (Revision != review.Revision) throw new InvalidOperationException("The document changed after orientation detection. Review it again.");
                return review.Review.ApplyResult(LoadDocument(bytes), captured, token).RequireValue().ToBytes();
            }, token, progress);
    }
}

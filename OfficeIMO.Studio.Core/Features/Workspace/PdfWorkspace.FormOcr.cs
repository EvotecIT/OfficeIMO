using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed record PdfWorkspaceFormOcrReview(PdfWorkspace Owner, long Revision, PdfFormOcrReview Review);

internal sealed partial class PdfWorkspace {
    internal async Task<PdfWorkspaceFormOcrReview> PrepareFormOcrAsync(IOcrEngine engine,
        PdfOcrMergeOptions options, CancellationToken token) {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(engine);
        ArgumentNullException.ThrowIfNull(options);
        var capturedOptions = options.Clone();
        long revision;
        byte[] bytes;
        await _operationGate.WaitAsync(token).ConfigureAwait(false);
        try {
            ThrowIfDisposed(); token.ThrowIfCancellationRequested();
            revision = Revision;
            // Workspace byte arrays are private immutable snapshots. Capture the exact pair without cloning/parsing on UI.
            bytes = _bytes;
        } finally { _operationGate.Release(); }
        var review = await RunNonDetachableCpuWorkAsync<OfficeIMO.Pdf.Ocr.PdfFormOcrReview>(
            () => new OfficeWorkflowRunner().PreparePdfFormOcrAsync(LoadDocument(bytes), engine, capturedOptions, cancellationToken: token), token).ConfigureAwait(false);
        token.ThrowIfCancellationRequested();
        ThrowIfDisposed();
        if (Revision != revision) throw new InvalidOperationException("The document changed during form recognition.");
        return new(this, revision, review);
    }

    internal Task ApplyFormOcrAsync(PdfWorkspaceFormOcrReview review,
        IReadOnlyDictionary<PdfFormOcrProposal, PdfFormFieldValue> accepted, CancellationToken token,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        if (!ReferenceEquals(review.Owner, this)) throw new ArgumentException("The recognition belongs to another document.", nameof(review));
        var snapshot = accepted.ToDictionary(pair => pair.Key, pair => PdfFormFieldValue.FromValues(pair.Value.Values));
        if (snapshot.Count == 0) throw new ArgumentException("Accept at least one reviewed value.", nameof(accepted));
        return MutateBytesAsync(PdfWorkspaceOperationKind.FormFill, "Applied reviewed form recognition", [], bytes => {
            if (Revision != review.Revision) throw new InvalidOperationException("The document changed after recognition. Recognize it again.");
            return review.Review.Apply(LoadDocument(bytes), snapshot, token).ToBytes();
        }, token, progress);
    }
}

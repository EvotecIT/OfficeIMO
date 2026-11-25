namespace OfficeIMO.Pdf;

/// <summary>Bounds the asset pass independently of the limits on parsing page content.</summary>
internal sealed class PdfAssetInspectionBudget {
    private readonly int _maximumWork;
    private long _used;

    internal PdfAssetInspectionBudget(int maximumWork, CancellationToken cancellationToken) {
        _maximumWork = maximumWork;
        CancellationToken = cancellationToken;
    }

    internal CancellationToken CancellationToken { get; }

    internal void Consume(long units) {
        CancellationToken.ThrowIfCancellationRequested();
        _used = checked(_used + units);
        if (_used > _maximumWork) throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts, _maximumWork, _used);
    }
}

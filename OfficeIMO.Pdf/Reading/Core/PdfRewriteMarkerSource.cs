namespace OfficeIMO.Pdf;

/// <summary>
/// Shares an unmodified object graph across a preflight's feature-specific rewrite checks.
/// The canonical reader can repair its object graph in memory, so repaired or encrypted inputs
/// retain the original-byte parse for conservative rewrite decisions.
/// </summary>
internal sealed class PdfRewriteMarkerSource {
    private readonly Lazy<(Dictionary<int, PdfIndirectObject> Map, string TrailerRaw)> _parsed;
    private readonly byte[] _pdf;
    private readonly PdfLoadOptions? _options;
    private readonly bool _usesDocumentSnapshot;
    private (Dictionary<int, PdfIndirectObject> Map, string TrailerRaw)? _cancellableParsed;

    internal PdfRewriteMarkerSource(byte[] pdf, PdfLoadOptions? options, PdfReadDocument? readDocument = null) {
        _pdf = pdf;
        _options = options;
        _usesDocumentSnapshot = readDocument is not null &&
            ReferenceEquals(readDocument.ReadOptions, options) &&
            !readDocument.Security.HasEncryption &&
            !readDocument.RepairReport.HasRepairs;
        _parsed = _usesDocumentSnapshot
                ? new Lazy<(Dictionary<int, PdfIndirectObject> Map, string TrailerRaw)>(
                    () => (readDocument!.Objects, readDocument.TrailerRaw))
                : new Lazy<(Dictionary<int, PdfIndirectObject> Map, string TrailerRaw)>(
                    () => PdfSyntax.ParseObjects(pdf, options));
    }

    internal (Dictionary<int, PdfIndirectObject> Map, string TrailerRaw) Parse(
        System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!cancellationToken.CanBeCanceled || _usesDocumentSnapshot || _parsed.IsValueCreated) return _parsed.Value;
        if (_cancellableParsed.HasValue) return _cancellableParsed.Value;
        var parsed = PdfSyntax.ParseObjects(_pdf, _options, out _, out _, cancellationToken);
        _cancellableParsed = parsed;
        return parsed;
    }
}

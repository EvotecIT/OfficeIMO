namespace OfficeIMO.Pdf;

/// <summary>
/// Shares an unmodified object graph across a preflight's feature-specific rewrite checks.
/// The canonical reader can repair its object graph in memory, so repaired or encrypted inputs
/// retain the original-byte parse for conservative rewrite decisions.
/// </summary>
internal sealed class PdfRewriteMarkerSource {
    private readonly Lazy<(Dictionary<int, PdfIndirectObject> Map, string TrailerRaw)> _parsed;

    internal PdfRewriteMarkerSource(byte[] pdf, PdfLoadOptions? options, PdfReadDocument? readDocument = null) {
        _parsed = readDocument is not null &&
            ReferenceEquals(readDocument.ReadOptions, options) &&
            !readDocument.Security.HasEncryption &&
            !readDocument.RepairReport.HasRepairs
                ? new Lazy<(Dictionary<int, PdfIndirectObject> Map, string TrailerRaw)>(
                    () => (readDocument.Objects, readDocument.TrailerRaw))
                : new Lazy<(Dictionary<int, PdfIndirectObject> Map, string TrailerRaw)>(
                    () => PdfSyntax.ParseObjects(pdf, options));
    }

    internal (Dictionary<int, PdfIndirectObject> Map, string TrailerRaw) Parse() => _parsed.Value;
}

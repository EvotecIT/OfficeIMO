namespace OfficeIMO.Pdf;

/// <summary>
/// Lazily shares one unmodified syntax parse across a preflight's feature-specific rewrite checks.
/// The canonical reader can repair its object graph in memory, so these checks must use their
/// original-byte parse to preserve the existing conservative rewrite decisions.
/// </summary>
internal sealed class PdfRewriteMarkerSource {
    private readonly Lazy<(Dictionary<int, PdfIndirectObject> Map, string TrailerRaw)> _parsed;

    internal PdfRewriteMarkerSource(byte[] pdf, PdfLoadOptions? options) {
        _parsed = new Lazy<(Dictionary<int, PdfIndirectObject> Map, string TrailerRaw)>(
            () => PdfSyntax.ParseObjects(pdf, options));
    }

    internal (Dictionary<int, PdfIndirectObject> Map, string TrailerRaw) Parse() => _parsed.Value;
}

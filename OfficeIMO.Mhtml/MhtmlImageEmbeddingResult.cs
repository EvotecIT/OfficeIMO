using OfficeIMO.Html;

namespace OfficeIMO.Mhtml;

/// <summary>
/// Editable HTML projection with archived image bytes embedded and resource losses reported.
/// Target adapters may add their own format-specific diagnostics when importing <see cref="HtmlConversionResult{T}.Value"/>.
/// </summary>
public sealed class MhtmlImageEmbeddingResult : HtmlConversionResult<HtmlConversionDocument> {
    internal MhtmlImageEmbeddingResult(HtmlConversionDocument document,
        IReadOnlyList<HtmlDiagnostic> diagnostics, int embeddedResourceCount, long embeddedResourceBytes)
        : base(document) {
        AddDiagnostics(diagnostics);
        EmbeddedResourceCount = embeddedResourceCount;
        EmbeddedResourceBytes = embeddedResourceBytes;
    }

    /// <summary>Number of distinct archived images embedded in the editable document.</summary>
    public int EmbeddedResourceCount { get; }

    /// <summary>Total decoded bytes across distinct embedded images.</summary>
    public long EmbeddedResourceBytes { get; }
}

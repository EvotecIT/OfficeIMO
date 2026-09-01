using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Pdf;

/// <summary>Caller-provider OCR enrichment for a loaded PDF document.</summary>
public sealed class PdfDocumentOcr {
    private readonly PdfDocument _document;

    internal PdfDocumentOcr(PdfDocument document) => _document = document;

    /// <summary>
    /// Renders the selected pages, invokes the caller-owned OCR provider, and merges accepted words
    /// into the same logical result contract returned by <see cref="PdfDocument.Read"/>.
    /// </summary>
    public Task<PdfOcrMergeResult> ReadAsync(
        IPdfOcrProvider provider,
        PdfOcrMergeOptions? options = null,
        CancellationToken cancellationToken = default) =>
        PdfOcr.RecognizeAndMergeAsync(
            _document.GetBytesForOperation(),
            provider,
            options,
            _document.ReadOptions,
            cancellationToken);
}

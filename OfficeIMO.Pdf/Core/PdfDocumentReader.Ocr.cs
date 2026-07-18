using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentReader {
    /// <summary>Runs external OCR through the engine-owned render and native-text merge contract.</summary>
    public Task<PdfOcrMergeResult> OcrAsync(
        IPdfOcrProvider provider,
        PdfOcrMergeOptions? options = null,
        PdfReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        return PdfOcr.RecognizeAndMergeAsync(_document.GetBytesForOperation(), provider, options, ResolveReadOptions(readOptions), cancellationToken);
    }
}

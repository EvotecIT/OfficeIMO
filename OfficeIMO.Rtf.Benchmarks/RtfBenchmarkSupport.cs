using OfficeIMO.Rtf.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Benchmarks;

internal static class RtfBenchmarkSupport {
    public static RtfPdfSaveOptions CreatePdfSaveOptions() {
        var pdfOptions = new PdfCore.PdfOptions();
        if (!pdfOptions.TryUseDefaultDocumentFontFallback(requireEmbeddedFont: true)) {
            throw new InvalidOperationException(
                "The RTF PDF benchmark requires one installed Unicode font from the OfficeIMO document fallback list.");
        }

        return new RtfPdfSaveOptions { PdfOptions = pdfOptions };
    }
}

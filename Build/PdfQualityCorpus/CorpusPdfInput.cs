using OfficeIMO.Pdf;

namespace OfficeIMO.PdfQualityCorpus;

/// <summary>Loads an independent PDF snapshot from in-memory corpus bytes.</summary>
internal static class CorpusPdfInput {
    internal static PdfDocument Load(byte[] bytes) {
        using var input = new MemoryStream(bytes, writable: false);
        return PdfDocument.Load(input);
    }
}

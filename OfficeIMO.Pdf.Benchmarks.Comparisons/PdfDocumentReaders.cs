using System.Text;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using OfficePdfDocument = OfficeIMO.Pdf.PdfDocument;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal enum PdfReaderEngine {
    OfficeIMO,
    PdfPig,
    IText
}

internal static class PdfDocumentReaders {
    internal static PdfReadObservation Read(PdfReaderEngine engine, byte[] pdf) => engine switch {
        PdfReaderEngine.OfficeIMO => ReadWithOfficeImo(pdf),
        PdfReaderEngine.PdfPig => PdfBenchmarkValidation.ReadWithPdfPig(pdf),
        PdfReaderEngine.IText => ReadWithIText(pdf),
        _ => throw new ArgumentOutOfRangeException(nameof(engine))
    };

    internal static string ExtractText(PdfReaderEngine engine, byte[] pdf) => engine switch {
        PdfReaderEngine.OfficeIMO => OfficePdfDocument.Open(pdf).Read.Text(),
        PdfReaderEngine.PdfPig => ExtractTextWithPdfPig(pdf),
        PdfReaderEngine.IText => ExtractTextWithIText(pdf),
        _ => throw new ArgumentOutOfRangeException(nameof(engine))
    };

    private static PdfReadObservation ReadWithOfficeImo(byte[] pdf) {
        global::OfficeIMO.Pdf.PdfReadDocument document = global::OfficeIMO.Pdf.PdfReadDocument.Open(pdf);
        string text = document.ExtractText();
        return PdfBenchmarkValidation.Observe(document.Pages.Count, text);
    }

    private static PdfReadObservation ReadWithIText(byte[] pdf) {
        string text = ExtractTextWithIText(pdf, out int pageCount);
        return PdfBenchmarkValidation.Observe(pageCount, text);
    }

    private static string ExtractTextWithIText(byte[] pdf) => ExtractTextWithIText(pdf, out _);

    private static string ExtractTextWithIText(byte[] pdf, out int pageCount) {
        using var input = new MemoryStream(pdf, writable: false);
        using var reader = new iText.Kernel.Pdf.PdfReader(input);
        using var document = new iText.Kernel.Pdf.PdfDocument(reader);
        var text = new StringBuilder();
        pageCount = document.GetNumberOfPages();
        for (int page = 1; page <= pageCount; page++) {
            text.Append(PdfTextExtractor.GetTextFromPage(
                document.GetPage(page),
                new LocationTextExtractionStrategy()));
            text.Append('\n');
        }

        return text.ToString();
    }

    private static string ExtractTextWithPdfPig(byte[] pdf) {
        using UglyToad.PdfPig.PdfDocument document = UglyToad.PdfPig.PdfDocument.Open(pdf);
        var text = new StringBuilder();
        foreach (var page in document.GetPages()) {
            text.Append(page.Text);
            text.Append('\n');
        }

        return text.ToString();
    }
}

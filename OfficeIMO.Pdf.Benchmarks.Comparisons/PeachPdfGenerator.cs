using PeachPDF;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class PeachPdfGenerator {
    internal static byte[] Generate(string html) {
        var config = new PdfGenerateConfig {
            PageSize = PeachPDF.PageSize.A4,
            PageOrientation = PeachPDF.PageOrientation.Portrait
        };
        var generator = new PdfGenerator();
        var document = generator.GeneratePdf(html, config).GetAwaiter().GetResult();
        using var output = new MemoryStream();
        document.Save(output);
        return output.ToArray();
    }
}

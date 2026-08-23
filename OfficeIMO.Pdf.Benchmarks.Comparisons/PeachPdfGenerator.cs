using PeachPDF;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class PeachPdfGenerator {
    internal static byte[] Generate(string html, PdfEmbeddedFontFamily? additionalFont = null) {
        var config = new PdfGenerateConfig {
            PageSize = PeachPDF.PageSize.A4,
            PageOrientation = PeachPDF.PageOrientation.Portrait,
            EnableTaggedPdf = true
        };
        var generator = new PdfGenerator();
        if (additionalFont != null) {
            using var fontStream = new MemoryStream(additionalFont.Regular, writable: false);
            generator.AddFontFromStream(fontStream);
        }
        var document = generator.GeneratePdf(html, config).GetAwaiter().GetResult();
        using var output = new MemoryStream();
        document.Save(output);
        return output.ToArray();
    }
}

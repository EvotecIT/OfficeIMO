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
            foreach (byte[] face in new[] { additionalFont.Regular, additionalFont.Bold, additionalFont.Italic, additionalFont.BoldItalic }.OfType<byte[]>()) {
                using var fontStream = new MemoryStream(face, writable: false);
                generator.AddFontFromStream(fontStream).GetAwaiter().GetResult();
            }
        }
        var document = generator.GeneratePdf(html, config).GetAwaiter().GetResult();
        using var output = new MemoryStream();
        document.Save(output);
        return output.ToArray();
    }
}

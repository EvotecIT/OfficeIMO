using iText.Html2pdf;
using iText.Kernel.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class ITextPdfHtmlGenerator {
    internal static byte[] Generate(string html) {
        using var output = new MemoryStream();
        var writer = new PdfWriter(output, new WriterProperties().SetCompressionLevel(6));
        var pdf = new iText.Kernel.Pdf.PdfDocument(writer);
        pdf.SetTagged();
        HtmlConverter.ConvertToPdfBytes(html, pdf, new ConverterProperties());
        return output.ToArray();
    }
}

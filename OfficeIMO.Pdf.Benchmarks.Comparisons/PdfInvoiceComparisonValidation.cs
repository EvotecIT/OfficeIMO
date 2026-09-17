using System.Text;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class PdfInvoiceComparisonValidation {
    internal static void Validate(byte[] bytes, InvoiceComparisonScenario scenario, string engine) {
        if (bytes.Length < 5 || !bytes.AsSpan(0, 4).SequenceEqual("%PDF"u8))
            throw new InvalidDataException($"{engine} did not produce a PDF header.");
        using var stream = new MemoryStream(bytes, writable: false);
        using UglyToad.PdfPig.PdfDocument document = UglyToad.PdfPig.PdfDocument.Open(stream);
        if (document.NumberOfPages != 2)
            throw new InvalidDataException($"{engine} produced {document.NumberOfPages} invoice pages; expected 2.");
        var text = new StringBuilder();
        foreach (var page in document.GetPages()) text.AppendLine(ContentOrderTextExtractor.GetText(page));
        string normalized = Normalize(text.ToString());
        foreach (string required in scenario.RequiredText) {
            if (!normalized.Contains(Normalize(required), StringComparison.Ordinal))
                throw new InvalidDataException($"{engine} omitted required invoice content: {required}");
        }
        ValidateLogoAspectRatio(document.GetPage(1), engine);
    }

    private static void ValidateLogoAspectRatio(UglyToad.PdfPig.Content.Page page, string engine) {
        const double expectedLogoAspectRatio = 5D;
        UglyToad.PdfPig.Content.IPdfImage? logo = page.GetImages()
            .Where(static image => !image.IsImageMask && image.WidthInSamples > 0 && image.HeightInSamples > 0)
            .OrderBy(image => Math.Abs(image.WidthInSamples / (double)image.HeightInSamples - expectedLogoAspectRatio))
            .FirstOrDefault();
        if (logo is null || Math.Abs(logo.WidthInSamples / (double)logo.HeightInSamples - expectedLogoAspectRatio) > 0.01D)
            throw new InvalidDataException($"{engine} omitted the expected 5:1 comparison logo.");

        double renderedAspectRatio = logo.BoundingBox.Width / logo.BoundingBox.Height;
        if (Math.Abs(renderedAspectRatio - expectedLogoAspectRatio) > 0.01D)
            throw new InvalidDataException(
                $"{engine} distorted the comparison logo to {renderedAspectRatio:0.###}:1; expected 5:1.");
    }

    private static string Normalize(string value) {
        var text = new StringBuilder(value.Length);
        foreach (char character in value.ToUpperInvariant())
            if (char.IsLetterOrDigit(character)) text.Append(character);
        return text.ToString();
    }
}

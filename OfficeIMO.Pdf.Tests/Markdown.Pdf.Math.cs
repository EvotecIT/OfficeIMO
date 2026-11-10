using System.Text;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class MarkdownPdfTests {
    [Fact]
    public void Markdown_SaveAsPdf_LaysOutTexMathOrKeepsVisibleSourceWithAWarning() {
        const string markdown = "Energy $E = mc^2$ holds.\n\n$$\n\\frac{a}{b}\n$$\n\n```math\n\\begin{unknownenv}x\\end{unknownenv}\n```";
        MarkdownDocument document = MarkdownDocument.Parse(markdown, new MarkdownLoadOptions { Math = true });

        PdfCore.PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(new MarkdownPdfExportOptions());
        string text = PdfCore.PdfReadDocument.Open(result.ToBytes()).ExtractText();

        Assert.Contains("Energy", text, StringComparison.Ordinal);
        Assert.Contains("unknownenv", text, StringComparison.Ordinal);
        if (result.Warnings.Any(warning => warning.Code == "math-font-unavailable")) {
            // Hosts without an installed math font keep every formula readable as source.
            Assert.Contains("\\frac{a}{b}", text, StringComparison.Ordinal);
            return;
        }

        Assert.Single(result.Warnings, warning => warning.Code == "math-unsupported");
        Assert.DoesNotContain("\\frac", text, StringComparison.Ordinal);
        Assert.DoesNotContain("mc^2", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_SaveAsPdf_DrawsSvgImagesAsVectorContent() {
        const string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"120\" height=\"40\"><rect width=\"120\" height=\"40\" fill=\"#2563eb\"/></svg>";
        string markdown = "![Blue bar](data:image/svg+xml;base64," + Convert.ToBase64String(Encoding.UTF8.GetBytes(svg)) + ")";

        PdfCore.PdfDocumentConversionResult result = MarkdownDocument.Parse(markdown).ConvertToPdfWithEvidence(new MarkdownPdfExportOptions());
        byte[] bytes = result.ToBytes();
        string pdf = Encoding.ASCII.GetString(bytes);

        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "unsupported-image");
        Assert.DoesNotContain("/Subtype /Image", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("Image unavailable", PdfCore.PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_SaveAsPdf_ShowsAPlaceholderForAnEmptySvgScene() {
        const string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"120\" height=\"40\"></svg>";
        string markdown = "![Empty](data:image/svg+xml;base64," + Convert.ToBase64String(Encoding.UTF8.GetBytes(svg)) + ")";

        PdfCore.PdfDocumentConversionResult result = MarkdownDocument.Parse(markdown).ConvertToPdfWithEvidence(new MarkdownPdfExportOptions());

        Assert.Contains(result.Warnings, warning => warning.Code == "unsupported-image");
        Assert.Contains("Image unavailable", PdfCore.PdfReadDocument.Open(result.ToBytes()).ExtractText(), StringComparison.Ordinal);
    }
}

using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class MarkdownPdfTableDefaultsTests {
    [Theory]
    [InlineData(null)]
    [InlineData(false)]
    [InlineData(true)]
    public void TablePresentationIsOptInAndAuthoredAlignmentIsPreserved(bool? applyTheme) {
        const string markdown = "| Default | Authored |\n| --- | ---: |\n| 12 | 12 |\n| 3456 | 3456 |\n";
        var options = new MarkdownPdfExportOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        };
        if (applyTheme.HasValue) options.ApplyDefaultTheme = applyTheme.Value;

        using var document = MarkdownDocument.Parse(markdown).ToPdf(options);
        byte[] bytes = document.ToBytes();
        string content = PdfTestContentText.ReadDecoded(bytes);
        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var words = pdf.GetPage(1).GetWords().ToArray();
        var shortValues = words.Where(word => word.Text == "12").OrderBy(word => word.BoundingBox.Left).ToArray();
        var longValues = words.Where(word => word.Text == "3456").OrderBy(word => word.BoundingBox.Left).ToArray();

        Assert.Equal(2, shortValues.Length);
        Assert.Equal(2, longValues.Length);
        Assert.InRange(Math.Abs(shortValues[0].BoundingBox.Left - longValues[0].BoundingBox.Left), 0D, .1D);
        Assert.InRange(Math.Abs(shortValues[1].BoundingBox.Right - longValues[1].BoundingBox.Right), 0D, .1D);
        Assert.Equal(applyTheme == true, content.Contains(" re f", StringComparison.Ordinal));
    }

    [Fact]
    public void ExplicitNumericAlignmentPolicyIsRetained() {
        const string markdown = "| Qty |\n| --- |\n| 12 |\n| 3456 |\n";
        var options = new MarkdownPdfExportOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
            Style = new MarkdownPdfStyle { TableStyle = new PdfCore.PdfTableStyle { RightAlignNumeric = true } }
        };

        using var document = MarkdownDocument.Parse(markdown).ToPdf(options);
        using var pdf = PdfPigDocument.Open(new MemoryStream(document.ToBytes()));
        var values = pdf.GetPage(1).GetWords().Where(word => word.Text is "12" or "3456").ToArray();

        Assert.Equal(2, values.Length);
        Assert.InRange(Math.Abs(values[0].BoundingBox.Right - values[1].BoundingBox.Right), 0D, .1D);
    }
}

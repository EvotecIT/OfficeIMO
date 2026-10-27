using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDecimalTabTests {
    [Theory]
    [InlineData(PdfTextWhitespaceMode.Preserve, '.', "1,234.56", "1234.56")]
    [InlineData(PdfTextWhitespaceMode.Preserve, ',', "1.234,56", "1234,56")]
    [InlineData(PdfTextWhitespaceMode.Preformatted, '.', "1,234.56", "1234.56")]
    [InlineData(PdfTextWhitespaceMode.Preformatted, ',', "1.234,56", "1234,56")]
    public void ExplicitDecimalSeparatorSurvivesStyleCopyAndIgnoresGrouping(PdfTextWhitespaceMode mode, char separator, string grouped, string control) {
        var style = new PdfParagraphStyle();
        style.AddTabStop(200, PdfTabAlignment.DecimalSeparator).DecimalSeparator = separator;
        var options = new PdfCreateOptions { PageWidth = 500, PageHeight = 300, MarginLeft = 20, MarginRight = 20,
            MarginTop = 20, MarginBottom = 20, DefaultFontSize = 12, TextWhitespaceMode = mode, DefaultParagraphStyle = style };
        PdfDocument document = PdfDocument.Create(options).Paragraph(p => p.Text("Grouped\t" + grouped))
            .Paragraph(p => p.Text("Control\t" + control));
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var anchors = pdf.GetPage(1).Letters.Where(letter => letter.Value == separator.ToString()).ToArray();
        Assert.Equal(2, anchors.Length);
        Assert.All(anchors, anchor => Assert.InRange(Math.Abs(anchor.StartBaseLine.X - 220), 0, .02));
    }
}

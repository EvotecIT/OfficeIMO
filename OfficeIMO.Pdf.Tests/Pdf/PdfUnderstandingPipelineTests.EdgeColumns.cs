using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfUnderstandingPipelineTests {
    [Fact]
    public void RightAlignedLeftColumnContinuationDoesNotBecomeCenteredFooter() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        var options = PdfUnderstandingPipelineOptions.Structured();
        options.GlyphDecoding = new FixedGlyphStage(new[] {
            new PdfTextSpan("LeftOne", "F1", 12, 50, 700, 230),
            new PdfTextSpan("LeftTwo", "F1", 12, 50, 650, 230),
            new PdfTextSpan("RightOne", "F1", 12, 320, 700, 230),
            new PdfTextSpan("RightTwo", "F1", 12, 320, 650, 230),
            new PdfTextSpan("Continuation", "F1", 12, 230, 450, 50)
        });
        var page = Assert.Single(Read(pdf, options).Pages).Analysis;
        Assert.Equal(new[] { "LeftOne", "LeftTwo", "Continuation", "RightOne", "RightTwo" },
            page.ReadingOrder.Select(static region => region.Text));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SeparatedColumnContinuationRetainsItsColumn(bool rightToLeft) {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        var options = PdfUnderstandingPipelineOptions.Structured();
        options.GlyphDecoding = new FixedGlyphStage(new[] {
            new PdfTextSpan("Left one", "F1", 12, 50, 700, 110),
            new PdfTextSpan("Left two", "F1", 12, 50, 650, 110),
            new PdfTextSpan("Right one", "F1", 12, 320, 700, 110),
            new PdfTextSpan("Right two", "F1", 12, 320, 650, 110),
            new PdfTextSpan("Left continuation", "F1", 12, 50, 450, 110)
        });
        var layout = new PdfTextLayoutOptions {
            ReadingDirection = rightToLeft ? PdfReadingDirection.RightToLeft : PdfReadingDirection.LeftToRight
        };
        var page = Assert.Single(Read(pdf, options, layoutOptions: layout).Pages).Analysis;
        Assert.Equal(rightToLeft
            ? new[] { "Right one", "Right two", "Left one", "Left two", "Left continuation" }
            : new[] { "Left one", "Left two", "Left continuation", "Right one", "Right two" },
            page.ReadingOrder.Select(static region => region.Text));
    }
}

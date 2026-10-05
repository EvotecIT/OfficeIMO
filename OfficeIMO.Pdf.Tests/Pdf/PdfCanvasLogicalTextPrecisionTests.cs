using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfCanvasLogicalTextPrecisionTests {
    [Theory]
    [InlineData(.2214321D)]
    [InlineData(.0044321D)]
    public void BoundedActualText_FractionalHeightPreservesAuthoredAdvance(double height) {
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .Canvas(canvas => canvas.ActualText("Logical marker", 10D, 20D, 50D, height,
                paint => paint.Text("PaintOnly", 10D, 20D, 50D, 12D)))
            .ToBytes();
        PdfReadPage page = Assert.Single(PdfReadDocument.Open(bytes).Pages);
        PdfTextSpan span = Assert.Single(page.GetTextSpans());
        Assert.Equal("Logical marker", span.Text);
        Assert.Equal(50D, span.Advance, 2);
        Assert.Equal(height, span.FontSize, 6);
        Assert.Equal(10D, span.X, 3);
    }
}

using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfLayoutDebugOverlayTests {
    [Fact]
    public void DebugOverlay_UsesSharedDrawingForWordsLinesRegionsAndReadingOrder() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Layout diagnostics"))
            .Paragraph(paragraph => paragraph.Text("First paragraph has several selectable words."))
            .Paragraph(paragraph => paragraph.Text("Second paragraph proves reading order."))
            .ToBytes();

        OfficeDrawing drawing = PdfDocument.Open(source).Read.LayoutDebugOverlay(1);
        string svg = PdfLayoutDebugOverlay.ToSvg(source, 1);
        byte[] png = PdfLayoutDebugOverlay.ToPng(source, 1, scale: 0.5D);

        Assert.Equal(612, drawing.Width);
        Assert.Equal(792, drawing.Height);
        Assert.NotEmpty(drawing.Shapes);
        Assert.NotEmpty(drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Contains("<rect", svg, StringComparison.Ordinal);
        Assert.Contains(">1</text>", svg, StringComparison.Ordinal);
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, png.Take(8).ToArray());
    }

    [Fact]
    public void DebugOverlay_EnforcesElementBudget() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Several words exceed one overlay element."))
            .ToBytes();

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfLayoutDebugOverlay.CreateDrawing(
                source,
                1,
                new PdfLayoutDebugOverlayOptions { MaxElements = 1 }));

        Assert.Equal(PdfReadLimitKind.InteractionRegions, exception.Kind);
        Assert.Equal(1, exception.Limit);
    }
}

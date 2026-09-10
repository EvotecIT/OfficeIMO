using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPositionedTextRenderingTests {
    [Theory]
    [InlineData(90)]
    [InlineData(180)]
    [InlineData(270)]
    public void PageRotationRetainsEverySourceGlyph(int rotation) {
        var document = OfficeIMO.Pdf.PdfDocument.Create(compose => compose.Page(page => page.Size(300, 400)
            .Content(content => content.Text("Needle and needle"))));
        // At 11px, a second raster sampling changes alpha coverage without dropping glyphs.
        // Supersampling separates missing ink from that pixel-phase effect and permits a tighter bound.
        OfficeRasterImage source = OfficeDrawingRasterRenderer.Render(document.Render.Drawing(1), scale: 4D);
        OfficeRasterImage rotated = OfficeDrawingRasterRenderer.Render(document.Pages.Rotate(rotation).Render.Drawing(1), scale: 4D);
        long sourceInk = source.GetPixels().Where((_, index) => index % 4 == 3).Sum(alpha => (long)alpha);
        long rotatedInk = rotated.GetPixels().Where((_, index) => index % 4 == 3).Sum(alpha => (long)alpha);
        Assert.True(sourceInk > 1000);
        Assert.InRange(rotatedInk, (long)(sourceInk * 0.98), (long)(sourceInk * 1.02));
    }
}

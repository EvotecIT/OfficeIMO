using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfReaderClipBoundsRegressionTests {
    [Fact]
    public void RenderPage_PreservesShapeThatBarelyExtendsBeyondPageClip() {
        OfficeColor blue = PdfColor.FromRgb(23, 58, 99).ToOfficeColor();
        OfficeShape background = OfficeShape.Rectangle(100D, 100D);
        background.FillColor = OfficeColor.White;
        background.StrokeWidth = 0D;
        OfficeShape foreground = OfficeShape.Rectangle(100D, 100D);
        foreground.FillColor = blue;
        foreground.StrokeWidth = 0D;

        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 100D,
                PageHeight = 100D,
                MarginLeft = 0D,
                MarginRight = 0D,
                MarginTop = 0D,
                MarginBottom = 0D,
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas
                .Shape(background, 0D, 0D)
                .Clip(0D, 0D, 100D, 100D, clipped => clipped
                    .Shape(foreground, 0.0005D, 0.0005D)))
            .ToBytes();

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));
        Assert.Equal(blue, raster.GetPixel(50, 50));
    }
}

using System;
using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingSvgTileSourcePixelTests {
    [Fact]
    public void FractionalImagePatternDecodesOneSourceWithinTheSharedBudget() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(16, 16, OfficeColor.Red));
        var tile = new OfficeDrawing(7.5, 4).AddImageWithInterpolation(png, "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(0, 0, 7.5, 4)), false);
        var drawing = new OfficeDrawing(30, 4).AddTilingPattern(tile,
            new OfficeImagePlacement(0, 0, 30, 4), 7.5, 4);
        // One 256-pixel source fits; decoding/charging it per placement does not.
        OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(drawing,
            new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 256 });
        for (int y = 0; y < actual.Height; y++)
            for (int x = 0; x < actual.Width; x++) Assert.Equal(OfficeColor.Red, actual.GetPixel(x, y));
    }

    [Theory]
    [InlineData(false, 30D)]
    [InlineData(true, 30D)]
    [InlineData(false, 29.25D)]
    [InlineData(true, 29.25D)]
    public void FractionalImagePatternSamplesSourceOnceThroughClip(bool multicolor, double clipWidth) {
        var source = new OfficeRasterImage(multicolor ? 2 : 1, 1, OfficeColor.Red);
        if (multicolor) source.SetPixel(1, 0, OfficeColor.Blue);
        byte[] png = OfficePngWriter.Encode(source);
        string href = "data:image/png;base64," + Convert.ToBase64String(png);
        string width = clipWidth.ToString(System.Globalization.CultureInfo.InvariantCulture);
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' width='60' height='26'>"
            + "<rect width='60' height='26' fill='white'/><defs>"
            + "<pattern id='p' x='8' y='8' width='7.5' height='4' patternUnits='userSpaceOnUse'>"
            + "<image width='7.5' height='4' preserveAspectRatio='none' image-rendering='optimizeSpeed'"
            + " style='image-rendering:pixelated' href='" + href + "'/></pattern></defs>"
            + "<rect x='8' y='8' width='" + width + "' height='4' fill='url(#p)'/></svg>");
        Assert.True(OfficeSvgDrawingReader.TryRead(svg, out OfficeDrawing? imported, out int unsupported));
        Assert.Equal(0, unsupported);

        var control = new OfficeDrawing(60, 26);
        OfficeShape background = OfficeShape.Rectangle(60, 26);
        background.FillColor = OfficeColor.White;
        background.StrokeColor = null;
        control.AddShape(background, 0, 0);
        for (int column = 0; column < 4; column++) {
            control.AddClippedImageWithInterpolation(png, "image/png",
                new OfficeImageProjection(new OfficeImagePlacement(8 + column * 7.5, 8, 7.5, 4)),
                false, 8, 8, OfficeClipPath.Rectangle(clipWidth, 4));
        }
        OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(imported!);
        Assert.Equal(OfficeDrawingRasterRenderer.Render(control).GetPixels(), actual.GetPixels());
        Assert.Equal(multicolor ? OfficeColor.Blue : OfficeColor.Red, actual.GetPixel(22, 9));
        Assert.Equal(clipWidth == 30D ? (multicolor ? OfficeColor.Blue : OfficeColor.Red) : OfficeColor.White,
            actual.GetPixel(37, 9));
        Assert.Equal(OfficeColor.White, actual.GetPixel(38, 9));
    }
}

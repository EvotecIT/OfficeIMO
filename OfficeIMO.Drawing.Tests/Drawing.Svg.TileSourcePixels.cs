using System;
using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingSvgTileSourcePixelTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void SvgTileSourcePixelsSurviveViewBoxScaling(bool pattern, bool interpolate) {
        var source = new OfficeRasterImage(4, 2);
        for (int y = 0; y < source.Height; y++) {
            for (int x = 0; x < source.Width; x++)
                source.SetPixel(x, y, OfficeColor.FromRgb((byte)(x * 31), (byte)(y * 29), (byte)((x + y) * 15)));
        }
        byte[] png = OfficePngWriter.Encode(source);
        string href = "data:image/png;base64," + Convert.ToBase64String(png);
        string sampling = interpolate ? "" : " image-rendering='optimizeSpeed' style='image-rendering:pixelated'";
        string Image(double x, double y) => "<image x='" + x + "' y='" + y
            + "' width='6' height='3' preserveAspectRatio='none' href='" + href + "'" + sampling + "/>";
        string content;
        if (pattern) {
            content = "<defs><pattern id='p' x='6' y='6' width='6' height='3' patternUnits='userSpaceOnUse'>"
                + Image(0, 0) + "</pattern></defs><rect x='6' y='6' width='15' height='7.5' fill='url(#p)'/>";
        } else {
            var tiles = new StringBuilder();
            for (int row = 0; row < 3; row++) {
                for (int column = 0; column < 3; column++) tiles.Append(Image(6 + column * 6, 6 + row * 3));
            }
            content = "<defs><clipPath id='c'><rect x='6' y='6' width='15' height='7.5'/></clipPath></defs>"
                + "<g clip-path='url(#c)'>" + tiles + "</g>";
        }
        byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='60' height='26' viewBox='0 0 45 19.5'>"
            + "<rect width='45' height='19.5' fill='white'/>" + content + "</svg>");
        Assert.True(OfficeSvgDrawingReader.TryRead(svg, out OfficeDrawing? imported, out int unsupported));
        Assert.Equal(0, unsupported);

        var control = new OfficeDrawing(60, 26);
        OfficeShape background = OfficeShape.Rectangle(60, 26);
        background.FillColor = OfficeColor.White;
        background.StrokeColor = null;
        control.AddShape(background, 0, 0);
        for (int row = 0; row < 3; row++) {
            for (int column = 0; column < 3; column++) {
                control.AddClippedImageWithInterpolation(png, "image/png",
                    new OfficeImageProjection(new OfficeImagePlacement(8 + column * 8, 8 + row * 4, 8, 4)),
                    interpolate, 8, 8, OfficeClipPath.Rectangle(20, 10));
            }
        }
        OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(imported!);
        OfficeRasterImage expected = OfficeDrawingRasterRenderer.Render(control);
        Assert.Equal(expected.GetPixels(), actual.GetPixels());
        if (!interpolate) {
            Assert.Equal(source.GetPixel(0, 0), actual.GetPixel(9, 9));
            Assert.Equal(source.GetPixel(3, 1), actual.GetPixel(15, 11));
            Assert.Equal(actual.GetPixel(9, 9), actual.GetPixel(25, 17));
        }
    }
}

using System;
using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingSvgRadialSpreadTests {
    [Theory]
    [InlineData("repeat", false, 0)]
    [InlineData("reflect", true, 0)]
    [InlineData("repeat", true, .1)]
    [InlineData("reflect", false, .1)]
    public void ShrinkingSpreadPreservesNegativeCyclesAndTransparentExterior(string spread, bool path, double endRadius) {
        string radius = endRadius.ToString(System.Globalization.CultureInfo.InvariantCulture);
        string geometry = path ? "<path d='M0,0H100V100H0Z' fill='url(#g)'/>" : "<rect width='100' height='100' fill='url(#g)'/>";
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='100' height='100'><defs><radialGradient id='g' fx='.3' fy='.5' fr='.25' cx='.7' cy='.5' r='" + radius + "' spreadMethod='" + spread + "'><stop offset='0' stop-color='red'/><stop offset='1' stop-color='blue'/></radialGradient></defs>" + geometry + "</svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out var drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        var image = OfficeDrawingRasterRenderer.Render(drawing!);
        double vx = .005 - .3, vy = .505 - .5, dr = endRadius - .25;
        double a = .4 * .4 - dr * dr, b = -2 * (vx * .4 + .25 * dr), c = vx * vx + vy * vy - .25 * .25;
        double t = (-b + Math.Sqrt(b * b - 4 * a * c)) / (2 * a);
        Assert.True(t < 0);
        double q = spread == "repeat" ? t - Math.Floor(t) : Math.Abs(t % 2);
        if (q > 1) q = 2 - q;
        Assert.InRange(Math.Abs(image.GetPixel(0, 50).B - 255 * q), 0, 1);
        Assert.Equal(0, image.GetPixel(95, 5).A);
        string saved = OfficeDrawingSvgExporter.ToSvg(drawing!, 1, OfficeSvgSizeUnit.Pixel);
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(saved), out var reopened, out unsupported));
        Assert.Equal(0, unsupported);
        Assert.Equal(image.GetPixel(0, 50), OfficeDrawingRasterRenderer.Render(reopened!).GetPixel(0, 50));
    }
    [Theory]
    [InlineData("repeat", "sRGB", 32, 128, 96)]
    [InlineData("reflect", "sRGB", 32, 128, 96)]
    [InlineData("repeat", "linearRGB", 99, 188, 165)]
    public void SvgBoundaryUsesOffsetWeightedColorAndAlpha(string spread, string interpolation, int red, int green, int blue) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='100' height='100'><defs><radialGradient id='g' fx='.75' fy='.5' fr='0' cx='.25' cy='.5' r='.5' spreadMethod='" + spread + "' color-interpolation='" + interpolation + "'><stop offset='0' stop-color='red' stop-opacity='.25'/><stop offset='.25' stop-color='#00ff00' stop-opacity='.5'/><stop offset='1' stop-color='blue'/></radialGradient></defs><rect width='100' height='100' fill='url(#g)'/></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out var drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        var pixel = OfficeDrawingRasterRenderer.Render(drawing!).GetPixel(90, 5);
        Assert.InRange(Math.Abs(pixel.R - red), 0, 1);
        Assert.InRange(Math.Abs(pixel.G - green), 0, 1);
        Assert.InRange(Math.Abs(pixel.B - blue), 0, 1);
        Assert.InRange(pixel.A, 167, 169);
        var exported = OfficeDrawingSvgExporter.ToSvg(drawing!, 1, OfficeSvgSizeUnit.Pixel);
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(exported), out var reopened, out unsupported));
        Assert.Equal(0, unsupported);
        var reopenedPixel = OfficeDrawingRasterRenderer.Render(reopened!).GetPixel(90, 5);
        Assert.InRange(Math.Abs(reopenedPixel.R - pixel.R), 0, 2);
        Assert.InRange(Math.Abs(reopenedPixel.A - pixel.A), 0, 2);
    }

    [Theory]
    [InlineData(OfficeGradientSpreadMode.Repeat)]
    [InlineData(OfficeGradientSpreadMode.Reflect)]
    public void ExplicitInteriorSpreadInterpolatesColorAndAlphaSeparately(OfficeGradientSpreadMode spread) {
        var gradient = new OfficeRadialGradient(.25, .5, 0, .25, .5, .5,
            new OfficeGradientStop(0, OfficeColor.FromRgba(255, 0, 0, 64)),
            new OfficeGradientStop(1, OfficeColor.FromRgba(0, 0, 255, 128))).WithSpreadMode(spread);
        var shape = OfficeShape.Rectangle(1, 1); shape.StrokeWidth = 0; shape.FillColor = null; shape.FillRadialGradient = gradient;
        var drawing = new OfficeDrawing(1, 1).AddShape(shape, 0, 0);
        Assert.Equal(OfficeColor.FromRgba(128, 0, 128, 96), OfficeDrawingRasterRenderer.Render(drawing).GetPixel(0, 0));
    }

}

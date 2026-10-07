using System;
using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingSvgShrinkingRadialTests {
    [Theory]
    [InlineData(false, false, 0)]
    [InlineData(false, true, 0)]
    [InlineData(true, false, 0)]
    [InlineData(true, true, 0)]
    [InlineData(false, false, .1)]
    [InlineData(true, false, .1)]
    public void ImportsShrinkingFieldsAndLeavesOutsideConeTransparent(bool userSpace, bool path, double endRadius) {
        string coordinates = userSpace
            ? "gradientUnits='userSpaceOnUse' fx='30' fy='50' fr='25' cx='70' cy='50' r='" + (endRadius * 100).ToString(System.Globalization.CultureInfo.InvariantCulture) + "'"
            : "fx='.3' fy='.5' fr='.25' cx='.7' cy='.5' r='" + endRadius.ToString(System.Globalization.CultureInfo.InvariantCulture) + "'";
        string geometry = path ? "<path d='M0,0H100V100H0Z' fill='url(#g)'/>" : "<rect width='100' height='100' fill='url(#g)'/>";
        var bytes = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='100' height='100'><defs><radialGradient id='g' " + coordinates + "><stop offset='0' stop-color='red'/><stop offset='1' stop-color='blue'/></radialGradient></defs>" + geometry + "</svg>");
        Assert.True(OfficeSvgDrawingReader.TryRead(bytes, out var drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        var image = OfficeDrawingRasterRenderer.Render(drawing!);
        Assert.Equal(0, image.GetPixel(95, 5).A);
        Assert.Equal(255, image.GetPixel(40, 50).A);
        // Along the horizontal diameter the physical root is approximately
        // (0.105 + 0.25) / (0.4 + 0.25 - endRadius).
        int blue = endRadius == 0 ? 139 : 165;
        Assert.InRange(image.GetPixel(40, 50).B, blue - 1, blue + 1);
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing!, 1, OfficeSvgSizeUnit.Pixel);
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out var reopened, out unsupported));
        Assert.Equal(0, unsupported);
        Assert.Equal(image.GetPixel(40, 50), OfficeDrawingRasterRenderer.Render(reopened!).GetPixel(40, 50));
        Assert.Equal(0, OfficeDrawingRasterRenderer.Render(reopened!).GetPixel(95, 5).A);
    }

    [Fact]
    public void TangentPointUsesTheLastPhysicalCircleAndTangentLineIsUnpainted() {
        var gradient = new OfficeRadialGradient(0, 0, 1, 1, 0, 0,
            new OfficeGradientStop(0, OfficeColor.Red), new OfficeGradientStop(1, OfficeColor.Blue));
        Assert.Equal(1D, gradient.SampleRatio(1, 0));
        Assert.True(double.IsNaN(gradient.SampleRatio(1, .1)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnpaintedConeIsTransparentForFillAndStroke(bool stroke) {
        var gradient = new OfficeRadialGradient(.3, .5, .25, .7, .5, 0,
            new OfficeGradientStop(0, OfficeColor.Red), new OfficeGradientStop(1, OfficeColor.Blue));
        var shape = OfficeShape.Rectangle(100, 100); shape.FillColor = null;
        if (stroke) { shape.StrokeWidth = 12; shape.StrokeRadialGradient = gradient; }
        else { shape.StrokeWidth = 0; shape.FillRadialGradient = gradient; }
        var drawing = new OfficeDrawing(100, 100).AddShape(shape, 0, 0);
        Assert.Equal(0, OfficeDrawingRasterRenderer.Render(drawing).GetPixel(95, 5).A);
        Assert.Equal(OfficeColor.White, OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.White).GetPixel(95, 5));
    }
}

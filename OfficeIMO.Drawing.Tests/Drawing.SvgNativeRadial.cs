using System;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingSvgNativeRadialTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void NativePaintKeepsDeclaredCanvasPlacementAndMarkerExtent(bool transformed, bool clipped) {
        var shape = OfficeShape.Path(100, 80, OfficePathCommand.MoveTo(-20, 30), OfficePathCommand.LineTo(120, 50));
        shape.FillColor = null; shape.StrokeWidth = 8;
        shape.StrokeRadialGradient = Field();
        shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 100, 120);
        if (transformed) shape.Transform = OfficeTransform.Translate(10, 15);
        if (clipped) shape.ClipPath = OfficeClipPath.Rectangle(100, 80);
        var xml = XElement.Parse(OfficeDrawingSvgExporter.ToSvg(new OfficeDrawing(300, 200).AddShape(shape, 30, 40)));
        XNamespace ns = "http://www.w3.org/2000/svg";
        var pattern = Assert.Single(xml.Descendants(ns + "pattern"));
        double x = transformed || clipped ? 0 : 30, y = transformed || clipped ? 0 : 40;
        double left = Number(pattern, "x"), top = Number(pattern, "y");
        Assert.True(left <= x - 240 && top <= y - 220);
        Assert.True(left + Number(pattern, "width") >= x + 340);
        Assert.True(top + Number(pattern, "height") >= y + 300);
        foreach (var gradient in pattern.Descendants(ns + "radialGradient")) {
            Assert.Equal("userSpaceOnUse", (string?)gradient.Attribute("gradientUnits"));
            // The radial unit circle maps to the declared 100x80 canvas, not the
            // thin path's bounding box or the arrowhead's own bounding box.
            Assert.Equal(FormattableString.Invariant($"matrix(30 0 0 24 {50 + x} {40 + y})"), (string?)gradient.Attribute("gradientTransform"));
        }
    }

    [Fact]
    public void NativePaintKeepsAlphaSeparateAndHonorsCancellationAndOutputBudget() {
        var shape = OfficeShape.Rectangle(100, 80); shape.FillRadialGradient = Field();
        var drawing = new OfficeDrawing(100, 80).AddShape(shape, 0, 0);
        var xml = XElement.Parse(OfficeDrawingSvgExporter.ToSvg(drawing)); XNamespace ns = "http://www.w3.org/2000/svg";
        var stops = xml.Descendants(ns + "radialGradient").Last().Elements(ns + "stop").ToArray();
        Assert.Equal("#808080", (string?)stops[0].Attribute("stop-color"));
        Assert.Equal("#404040", (string?)stops[1].Attribute("stop-color"));
        Assert.All(xml.Descendants(ns + "stop"), stop => Assert.Null(stop.Attribute("stop-opacity")));
        Assert.Throws<OperationCanceledException>(() => OfficeDrawingSvgExporter.ToSvg(drawing, 1, OfficeSvgSizeUnit.Pixel, null, null, new CancellationToken(true)));
        Assert.Throws<OfficeImageExportBatchLimitException>(() => OfficeDrawingSvgExporter.ToSvgBytes(drawing, 1, OfficeSvgSizeUnit.Pixel, null, null, 1000, CancellationToken.None));
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeAlphaInterpolationDoesNotChangeWithShapeDispatch(bool path) {
        var shape = path ? OfficeShape.Path(100, 80, OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(100, 0),
            OfficePathCommand.LineTo(100, 80), OfficePathCommand.LineTo(0, 80), OfficePathCommand.Close()) : OfficeShape.Rectangle(100, 80);
        shape.FillRadialGradient = Field(); shape.StrokeWidth = 0;
        var image = OfficeDrawingRasterRenderer.Render(new OfficeDrawing(100, 80).AddShape(shape, 0, 0), background: OfficeColor.White);
        // Along the center axis, original t=(focus-x)/(focus-center+radius).
        // x=.605 gives t approximately .494, alpha=96, RGB=(129,0,126).
        var pixel = image.GetPixel(60, 40);
        Assert.InRange((int)pixel.R, 207, 209);
        Assert.InRange((int)pixel.G, 158, 160);
        Assert.InRange((int)pixel.B, 205, 207);
    }

    private static double Number(XElement e, string name) => double.Parse((string)e.Attribute(name)!, CultureInfo.InvariantCulture);
    private static OfficeRadialGradient Field() => new OfficeRadialGradient(1, .5, 0, .5, .5, .3,
        new OfficeGradientStop(0, OfficeColor.FromRgba(255, 0, 0, 64)),
        new OfficeGradientStop(1, OfficeColor.FromRgba(0, 0, 255, 128))).WithFirstPadIntersection();
}

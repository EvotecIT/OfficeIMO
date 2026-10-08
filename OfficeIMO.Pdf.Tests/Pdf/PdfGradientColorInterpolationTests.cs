using System;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfGradientColorInterpolationTests {
    [Fact]
    public void RadialShadingCacheDistinguishesColorInterpolationSpaces() {
        var first = OfficeShape.Rectangle(80, 40); first.StrokeWidth = 0;
        first.FillRadialGradient = OfficeRadialGradient.Centered(OfficeColor.Red, OfficeColor.Blue);
        var second = OfficeShape.Rectangle(80, 40); second.StrokeWidth = 0;
        second.FillRadialGradient = first.FillRadialGradient.WithColorInterpolation(OfficeGradientColorInterpolation.LinearRgb);
        byte[] pdf = PdfDocument.Create(new PdfOptions { PageWidth = 100, PageHeight = 100,
            MarginTop = 0, MarginLeft = 0, MarginRight = 0, MarginBottom = 0 })
            .Compose(c => c.Page(p => p.Content(content => content.Shape(first).Shape(second)))).ToBytes();
        var raster = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(pdf).Pages[0].ToDrawing());
        var srgb = raster.GetPixel(60, 20); var linear = raster.GetPixel(60, 60);
        Assert.InRange(srgb.R, 122, 127);
        Assert.InRange(linear.R, 183, 188);
        Assert.InRange(linear.B, 187, 193);
        Assert.Equal(OfficeGradientColorInterpolation.Srgb, first.FillRadialGradient.ColorInterpolation);
    }
}

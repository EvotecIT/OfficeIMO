using System;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsBrushPdfTests {
    [Theory]
    [InlineData("visual-Tile", false)]
    [InlineData("visual-FlipX", true)]
    [InlineData("image-Tile", false)]
    [InlineData("image-FlipXY", true)]
    public void PdfPreservesTilePlacementAndFlipParity(string scenario, bool flip) {
        byte[] pdf = XpsBrushFixtures.Create(scenario).ToPdf();
        var drawing = PdfDocument.Load(pdf).Render.Drawing(1);
        var pixels = OfficeDrawingRasterRenderer.Render(drawing, 96D / 72D);
        Assert.True(pixels.GetPixel(15, 15).R > 200);
        Assert.True(pixels.GetPixel(25, 15).B > 200);
        Assert.True(flip ? pixels.GetPixel(35, 15).B > 200 : pixels.GetPixel(35, 15).R > 200);
        if (scenario.StartsWith("image", StringComparison.Ordinal))
            Assert.Contains("/Interpolate true", Encoding.ASCII.GetString(pdf));
    }

    [Theory]
    [InlineData("mask")]
    [InlineData("mask-transformed")]
    public void PdfPreservesAlphaMaskCoordinates(string scenario) {
        byte[] pdf = XpsBrushFixtures.Create(scenario).ToPdf();
        Assert.Contains("/S /Alpha", Encoding.ASCII.GetString(pdf));
        var pixels = OfficeDrawingRasterRenderer.Render(PdfDocument.Load(pdf).Render.Drawing(1), 96D / 72D);
        var low = pixels.GetPixel(10, 40); var high = pixels.GetPixel(90, 40);
        Assert.True(high.R > 200);
        Assert.True(low.A < high.A || low.G > high.G);
    }
}

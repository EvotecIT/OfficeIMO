using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioSvgPreviewGraphicsQualityTests {
    [Theory]
    [InlineData("stroke-dasharray='0 6' stroke-dashoffset='2' stroke-linecap='round'")]
    [InlineData("stroke-dasharray='6 0 4 3' stroke-dashoffset='-3'")]
    [InlineData("stroke-miterlimit='1' stroke-linejoin='miter'")]
    public void EmbeddedPreviewPreservesDashPhaseZeroEntriesAndMiterLimit(string attributes) {
        string svg = $"<svg xmlns='http://www.w3.org/2000/svg' width='80' height='60'><g transform='scale(2)' {attributes}><path d='M5 25 L20 5 L35 25' fill='none' stroke='black' stroke-width='2'/></g></svg>";
        byte[] data = Encoding.UTF8.GetBytes(svg);
        Assert.True(VisioSvgPreviewRasterizer.TryRasterize(data, out OfficeRasterImage? image));
        Assert.True(OfficeSvgDrawingReader.TryRead(data, out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        byte[] expected = OfficeDrawingRasterRenderer.Render(drawing!).GetPixels(), actual = image!.GetPixels();
        Assert.Equal(expected.Length, actual.Length);
        // Equivalent local coordinates can differ by one alpha level after inverse projection.
        for (int i = 0; i < expected.Length; i++) Assert.InRange(System.Math.Abs(expected[i] - actual[i]), 0, 1);
    }

    [Theory]
    [InlineData("M10 30 L40 30 L70 30")]
    [InlineData("M10 30 L70 30 M40 10 L40 50")]
    public void EmbeddedPreviewKeepsOneOpacityForTheWholeStroke(string path) {
        string svg = $"<svg xmlns='http://www.w3.org/2000/svg' width='80' height='60'><path d='{path}' fill='none' stroke='black' stroke-opacity='.5' stroke-width='4' stroke-linecap='round'/></svg>";
        Assert.True(VisioSvgPreviewRasterizer.TryRasterize(Encoding.UTF8.GetBytes(svg), out OfficeRasterImage? image));
        byte[] pixels = image!.GetPixels();
        Assert.Contains(Enumerable.Range(0, pixels.Length / 4).Select(i => pixels[i * 4 + 3]), alpha => alpha == 128);
        Assert.All(Enumerable.Range(0, pixels.Length / 4), i => Assert.InRange(pixels[i * 4 + 3], (byte)0, (byte)128));
    }

    [Theory]
    [InlineData("butt")]
    [InlineData("round")]
    [InlineData("square")]
    public void EmbeddedPreviewAndSharedSvgReaderAgreeOnFractionalStrokes(string cap) {
        string svg = $"<svg xmlns='http://www.w3.org/2000/svg' width='80' height='60'><path d='M10 30 L40 10 L70 30' fill='none' stroke='black' stroke-width='.25' stroke-linecap='{cap}' stroke-linejoin='bevel'/></svg>";
        byte[] data = Encoding.UTF8.GetBytes(svg);
        Assert.True(VisioSvgPreviewRasterizer.TryRasterize(data, out OfficeRasterImage? image));
        Assert.True(OfficeSvgDrawingReader.TryRead(data, out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.Equal(OfficeDrawingRasterRenderer.Render(drawing!).GetPixels(), image!.GetPixels());
    }
}

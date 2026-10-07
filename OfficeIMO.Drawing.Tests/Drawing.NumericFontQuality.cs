using System;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingNumericFontQualityTests {
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void SimulatedBoldKeepsTranslucentGlyphOpacity(int placement) {
        var image = new OfficeRasterImage(120, 80);
        var canvas = new OfficeRasterCanvas(image, font: OfficeTrueTypeFont.TryLoad(FontWithAdvance(700)));
        var color = new OfficeColor(20, 40, 60, 128);
        if (placement == 0) canvas.DrawText("A", 10, 10, 100, 60, color, 32, style: OfficeFontStyle.Bold);
        else if (placement == 1) canvas.DrawTextLine("A", 20, 10, 32, color, bold: true, rotationDegrees: 25, rotationCenterX: 20, rotationCenterY: 10);
        else canvas.DrawTextLineTransformed("A", 20, 10, 32, color, new OfficeTransform(1, 0, .25, 1, 0, 0), bold: true);
        byte[] pixels = image.GetPixels();
        Assert.Contains(Enumerable.Range(0, pixels.Length / 4).Select(i => pixels[i * 4 + 3]), alpha => alpha > 0);
        Assert.All(Enumerable.Range(0, pixels.Length / 4), i => Assert.InRange(pixels[i * 4 + 3], (byte)0, (byte)128));
    }

    [Fact]
    public void FontCopiesAndCacheKeysRetainExactFaceAttributes() {
        var face = new OfficeFontFaceDescriptor(600, 75D, OfficeFontSlant.Oblique, 20D);
        var font = new OfficeFontInfo("Scoped", 20, face, OfficeFontStyle.Underline);
        Assert.Equal(face, font.WithSize(40).WithFamilyName("Alias").Face);
        Assert.Equal(face, font.WithStyle(font.Style | OfficeFontStyle.Strikethrough).Face);
        Assert.NotEqual(font, font.WithFace(new OfficeFontFaceDescriptor(700, 75D, OfficeFontSlant.Oblique, 20D)));
        _ = default(OfficeFontInfo).GetHashCode();
    }

    [Fact]
    public void MeasurementAndPaintingChooseTheSameRegisteredNumericFace() {
        byte[] regular = FontWithAdvance(500), semibold = FontWithAdvance(1000);
        var fonts = new OfficeFontFaceCollection().Add("Scoped", regular, new OfficeFontFaceDescriptor(400))
            .Add("Scoped", semibold, new OfficeFontFaceDescriptor(600));
        var image = new OfficeRasterImage(100, 60);
        var canvas = new OfficeRasterCanvas(image, fonts: fonts);
        var normalFont = new OfficeFontInfo("Scoped", 20, new OfficeFontFaceDescriptor(400));
        var semiboldFont = new OfficeFontInfo("Scoped", 20, new OfficeFontFaceDescriptor(600));
        double regularWidth = OfficeTrueTypeFont.TryLoad(regular)!.Measure("A", 20);
        double semiboldWidth = OfficeTrueTypeFont.TryLoad(semibold)!.Measure("A", 20);
        Assert.NotEqual(regularWidth, semiboldWidth);
        Assert.Equal(regularWidth, canvas.MeasureText("A", normalFont));
        Assert.Equal(semiboldWidth, canvas.MeasureText("A", semiboldFont));
        Assert.Equal(regularWidth, canvas.MeasureText("A", normalFont));
        canvas.DrawText("A", 0, 0, 100, 60, OfficeColor.Black, semiboldFont);
        var expected = new OfficeRasterImage(100, 60);
        new OfficeRasterCanvas(expected, font: OfficeTrueTypeFont.TryLoad(semibold)).DrawText("A", 0, 0, 100, 60, OfficeColor.Black, 20);
        Assert.Equal(expected.GetPixels(), image.GetPixels());
    }

    [Theory]
    [InlineData(100)]
    [InlineData(500)]
    [InlineData(600)]
    [InlineData(900)]
    public void SvgRoundTripKeepsNumericFaceAttributes(int weight) {
        var face = new OfficeFontFaceDescriptor(weight, 75D, OfficeFontSlant.Oblique, 20D);
        var drawing = new OfficeDrawing(200, 90);
        drawing.AddText("A", 10, 10, 180, 60, new OfficeFontInfo("Scoped", 20, face), OfficeColor.Black);
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);
        Assert.Contains($"font-weight=\"{weight}\"", svg);
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out var imported, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.Equal(face, TextElements(imported!).Single().Font.Face);
    }

    [Theory]
    [InlineData(900, "lighter", 700)]
    [InlineData(700, "lighter", 400)]
    [InlineData(400, "lighter", 100)]
    [InlineData(300, "bolder", 400)]
    [InlineData(400, "bolder", 700)]
    [InlineData(700, "bolder", 900)]
    [InlineData(950, "bolder", 950)]
    [InlineData(50, "lighter", 50)]
    public void RelativeSvgWeightsUseTheInheritedNumericWeight(int inherited, string relative, int expected) {
        string svg = $"<svg xmlns='http://www.w3.org/2000/svg' width='200' height='90'><g font-weight='{inherited}'><text x='10' y='50' font-size='20' font-weight='{relative}'>A</text></g></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out var drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.Equal(expected, drawing!.Elements.OfType<OfficeDrawingText>().Single().Font.Face.Weight);
    }

    private static byte[] FontWithAdvance(int advance) {
        byte[] font = ManagedTextShapingTestAssets.CreateFont('A');
        int count = font[4] * 256 + font[5];
        for (int i = 0; i < count; i++) {
            int table = 12 + i * 16;
            if (Encoding.ASCII.GetString(font, table, 4) != "hmtx") continue;
            int offset = (font[table + 8] << 24) | (font[table + 9] << 16) | (font[table + 10] << 8) | font[table + 11];
            font[offset] = (byte)(advance >> 8); font[offset + 1] = (byte)advance;
            return font;
        }
        throw new InvalidOperationException("The generated font fixture has no hmtx table.");
    }

    private static System.Collections.Generic.IEnumerable<OfficeDrawingText> TextElements(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            if (element is OfficeDrawingText text) yield return text;
            else if (element is OfficeDrawingGroup group) foreach (OfficeDrawingText child in TextElements(group.Drawing)) yield return child;
            else if (element is OfficeDrawingEffectGroup effect) foreach (OfficeDrawingText child in TextElements(effect.Drawing)) yield return child;
        }
    }
}

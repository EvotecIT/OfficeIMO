using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingTextInkBoundsTests {
    [Fact]
    public void InkMeasurementExcludesEmptyLayoutSpaceAndTransformsOutlinePoints() {
        var fonts = new OfficeFontFaceCollection().Add("Ink", ManagedTextShapingTestAssets.CreateFont('A'));
        var canvas = new OfficeRasterCanvas(new OfficeRasterImage(100, 100), fonts: fonts);
        var font = new OfficeFontInfo("Ink", 20);
        var buffer = canvas.MeasurePositionedTextBounds("A", 2, 10, 200, 200, 20, font, 10,
            OfficeTextAlignment.Left, OfficeTextFeatureSettings.Default, "normal", 20,
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None);
        var ink = canvas.MeasurePositionedTextBounds("A", 2, 10, 200, 200, 20, font, 10,
            OfficeTextAlignment.Left, OfficeTextFeatureSettings.Default, "normal", 20,
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None, inkOnly: true,
            inkTransform: OfficeTransform.RotateDegrees(90).Then(OfficeTransform.Translate(100, 0)));
        Assert.True(ink.IsMeasured && ink.HasInk);
        Assert.Equal(202, buffer.Right, 6);
        Assert.Equal(70, ink.Left, 6); Assert.Equal(84, ink.Right, 6);
        Assert.Equal(2, ink.Top, 6); Assert.Equal(10, ink.Bottom, 6);
    }

    [Theory]
    [InlineData(OfficeFontStyle.Regular, OfficeTextDecorationStyle.None)]
    [InlineData(OfficeFontStyle.Bold | OfficeFontStyle.Italic, OfficeTextDecorationStyle.Wavy)]
    public void InkBoundsContainActualPositionedTextAndDecorationPaint(OfficeFontStyle style, OfficeTextDecorationStyle decoration) {
        var fonts = new OfficeFontFaceCollection().Add("Ink", ManagedTextShapingTestAssets.CreateFont('A'));
        var image = new OfficeRasterImage(80, 80);
        var canvas = new OfficeRasterCanvas(image, fonts: fonts);
        var ink = canvas.MeasurePositionedTextBounds("A", 20, 20, 40, 40, 20, new OfficeFontInfo("Ink", 20, style), 10,
            OfficeTextAlignment.Left, OfficeTextFeatureSettings.Default, "normal", 20, decoration,
            OfficeTextDecorationStyle.None, inkOnly: true);
        canvas.DrawPositionedText("A", 20, 20, 40, 40, OfficeColor.Black, 20, OfficeTextAlignment.Left,
            style, "Ink", 10, decoration, OfficeTextDecorationStyle.None, baselineFontSize: 20);
        int painted = 0;
        for (int y = 0; y < image.Height; y++) for (int x = 0; x < image.Width; x++) {
            if (image.GetPixel(x, y).A == 0) continue;
            painted++;
            Assert.InRange(x + .5D, ink.Left - 1D, ink.Right + 1D);
            Assert.InRange(y + .5D, ink.Top - 1D, ink.Bottom + 1D);
        }
        Assert.True(painted > 0);
        Assert.True(ink.IsMeasured && ink.HasInk);
    }

    [Fact]
    public void ColorInkUsesPaintedLayersInsteadOfUnpaintedBaseGlyph() {
        var fonts = new OfficeFontFaceCollection().Add("Color Ink", ManagedTextShapingTestAssets.CreateColorFont('A', baseGlyphHeight: 2000));
        var canvas = new OfficeRasterCanvas(new OfficeRasterImage(100, 100), fonts: fonts);
        var font = new OfficeFontInfo("Color Ink", 20);
        var buffer = canvas.MeasurePositionedTextBounds("A", 0, 30, 20, 20, 20, font, 10,
            OfficeTextAlignment.Left, OfficeTextFeatureSettings.Default, "normal", 20,
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None);
        var ink = canvas.MeasurePositionedTextBounds("A", 0, 30, 20, 20, 20, font, 10,
            OfficeTextAlignment.Left, OfficeTextFeatureSettings.Default, "normal", 20,
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None, inkOnly: true);
        Assert.True(ink.IsMeasured && ink.HasInk);
        Assert.True(ink.Top > buffer.Top);
        Assert.Equal(36, ink.Top, 6);
    }
}

using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingTextInkBoundsTests {
    [Fact]
    public void EmptyEnabledClipAxisRemovesTextContours() {
        var clip = new OfficeTextInkClip(0, 0, 0, 20, true, false, OfficeTransform.Identity);
        var triangle = new List<OfficePoint> { new(0, 0), new(10, 0), new(0, 10) };
        bool clipped = false; long budget = 100;
        Assert.Empty(clip.Apply(triangle, ref clipped, ref budget, default));
        Assert.True(clipped);
    }

    [Fact]
    public void ClipDoesNotInventInkWhereOnlyTheContourBoxIntersects() {
        var clip = new OfficeTextInkClip(6, 6, 2, 2, true, true, OfficeTransform.Identity);
        var triangle = new List<OfficePoint> { new(0, 0), new(10, 0), new(0, 10) };
        bool clipped = false; long budget = 100;
        Assert.Empty(clip.Apply(triangle, ref clipped, ref budget, default));
        Assert.True(clipped);
    }

    [Fact]
    public void ClipWorkLimitAndCancellationRejectIncompleteMeasurement() {
        var clip = new OfficeTextInkClip(0, 0, 5, 5, true, true, OfficeTransform.Identity);
        var triangle = new List<OfficePoint> { new(0, 0), new(10, 0), new(0, 10) };
        bool clipped = false; long budget = 2;
        Assert.Throws<NotSupportedException>(() => clip.Apply(triangle, ref clipped, ref budget, default));
        budget = 100;
        using var cancellation = new System.Threading.CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => clip.Apply(triangle, ref clipped, ref budget, cancellation.Token));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RectangularClippingMatchesPaintedTextAndIgnoresDisabledAxes(bool rotated) {
        var fonts = new OfficeFontFaceCollection().Add("Ink", ManagedTextShapingTestAssets.CreateFont('A'));
        var image = new OfficeRasterImage(80, 80);
        var canvas = new OfficeRasterCanvas(image, fonts: fonts);
        OfficeTransform transform = rotated ? OfficeTransform.RotateDegrees(30, 40, 40) : OfficeTransform.Identity;
        var clip = new OfficeTextInkClip(22, 0, 4, 1, horizontal: true, vertical: false, transform.Invert());
        var ink = canvas.MeasurePositionedTextBounds("A", 20, 20, 40, 40, 20, new OfficeFontInfo("Ink", 20), 10,
            OfficeTextAlignment.Left, OfficeTextFeatureSettings.Default, "normal", 20,
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None, inkOnly: true,
            inkTransform: transform, inkClips: new[] { clip });
        var expected = transform.TransformRectangleBounds(22, 26, 4, 14);
        Assert.True(ink.IsMeasured && ink.HasInk && ink.IsClipped);
        Assert.Equal(expected.Left, ink.Left, 6); Assert.Equal(expected.Right, ink.Right, 6);
        Assert.Equal(expected.Top, ink.Top, 6); Assert.Equal(expected.Bottom, ink.Bottom, 6);
        if (!rotated) {
            using (canvas.PushClipRectangle(22, 0, 4, 80))
                canvas.DrawPositionedText("A", 20, 20, 40, 40, OfficeColor.Black, 20,
                    OfficeTextAlignment.Left, OfficeFontStyle.Regular, "Ink", 10,
                    OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None, baselineFontSize: 20);
            int painted = 0;
            for (int y = 0; y < 80; y++) for (int x = 0; x < 80; x++) {
                if (image.GetPixel(x, y).A == 0) continue;
                painted++; Assert.InRange(x + .5D, ink.Left, ink.Right); Assert.InRange(y + .5D, ink.Top, ink.Bottom);
            }
            Assert.Equal(56, painted);
        }
    }

    [Fact]
    public void DisjointNestedClipsProduceNoTextBounds() {
        var fonts = new OfficeFontFaceCollection().Add("Ink", ManagedTextShapingTestAssets.CreateFont('A'));
        var canvas = new OfficeRasterCanvas(new OfficeRasterImage(80, 80), fonts: fonts);
        var ink = canvas.MeasurePositionedTextBounds("A", 20, 20, 40, 40, 20, new OfficeFontInfo("Ink", 20), 10,
            OfficeTextAlignment.Left, OfficeTextFeatureSettings.Default, "normal", 20,
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None, inkOnly: true,
            inkClips: new[] { new OfficeTextInkClip(22, 0, 4, 80, true, true, OfficeTransform.Identity),
                new OfficeTextInkClip(30, 0, 10, 80, true, true, OfficeTransform.Identity) });
        Assert.True(ink.IsMeasured && ink.IsClipped);
        Assert.False(ink.HasInk);
    }

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

using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingFilledInkBoundsTests {
    [Theory]
    [InlineData(OfficeFillRule.NonZero, true)]
    [InlineData(OfficeFillRule.EvenOdd, false)]
    public void BoundsAndRasterAgreeOnRepeatedContourFill(OfficeFillRule rule, bool filled) {
        var contours = new List<List<OfficePoint>> { Rectangle(2, 3, 12, 13), Rectangle(2, 3, 12, 13) };
        var image = new OfficeRasterImage(20, 20); var canvas = new OfficeRasterCanvas(image);
        var bounds = canvas.MeasureFilledContourBounds(contours, rule);
        canvas.FillContourPaint(contours, rule, (_, _) => OfficeColor.Black);
        Assert.True(bounds.IsMeasured); Assert.Equal(filled, bounds.HasInk);
        Assert.Equal(filled, image.GetPixel(5, 5).A > 0);
        if (filled) { Assert.Equal(2, bounds.Left); Assert.Equal(13, bounds.Bottom); }
    }

    [Fact]
    public void OppositeWindingRemovesCancelledExtents() {
        var removed = Rectangle(0, 0, 10, 20); removed.Reverse();
        var contours = new List<List<OfficePoint>> { Rectangle(0, 0, 20, 20), removed };
        var image = new OfficeRasterImage(30, 30); var canvas = new OfficeRasterCanvas(image);
        var bounds = canvas.MeasureFilledContourBounds(contours, OfficeFillRule.NonZero);
        canvas.FillContourPaint(contours, OfficeFillRule.NonZero, (_, _) => OfficeColor.Black);
        Assert.True(bounds.IsMeasured && bounds.HasInk);
        Assert.Equal(10, bounds.Left); Assert.Equal(20, bounds.Right);
        Assert.Equal(0, image.GetPixel(5, 5).A); Assert.Equal(255, image.GetPixel(15, 5).A);
    }

    [Fact]
    public void ClipInsideGlyphHoleHasNoFilledInk() {
        var hole = Rectangle(5, 5, 15, 15); hole.Reverse();
        var contours = new List<List<OfficePoint>> { Rectangle(0, 0, 20, 20), hole };
        var clip = new OfficeTextInkClip(6, 6, 8, 8, true, true, OfficeTransform.Identity);
        bool cropped = false; long work = 1000;
        for (int i = 0; i < contours.Count; i++) contours[i] = clip.Apply(contours[i], ref cropped, ref work, default);
        var image = new OfficeRasterImage(20, 20); var canvas = new OfficeRasterCanvas(image);
        var bounds = canvas.MeasureFilledContourBounds(contours, OfficeFillRule.NonZero);
        canvas.FillContourPaint(contours, OfficeFillRule.NonZero, (_, _) => OfficeColor.Black);
        Assert.True(cropped && bounds.IsMeasured); Assert.False(bounds.HasInk);
        for (int y = 0; y < 20; y++) for (int x = 0; x < 20; x++) Assert.Equal(0, image.GetPixel(x, y).A);
    }

    [Fact]
    public void SelfCrossingContourHasInkDespiteZeroSignedArea() {
        var contours = new List<List<OfficePoint>> { new() { new(0, 0), new(20, 20), new(0, 20), new(20, 0) } };
        var canvas = new OfficeRasterCanvas(new OfficeRasterImage(20, 20));
        var bounds = canvas.MeasureFilledContourBounds(contours, OfficeFillRule.NonZero);
        Assert.True(bounds.IsMeasured && bounds.HasInk);
        Assert.Equal(0, bounds.Left); Assert.Equal(20, bounds.Right);
        Assert.Equal(0, bounds.Top); Assert.Equal(20, bounds.Bottom);
    }

    [Fact]
    public void ExcessiveGeometryIsUnmeasuredAndCancellationIsHonored() {
        var points = Enumerable.Range(0, 5000).Select(i => new OfficePoint(i, i % 2)).ToList();
        var canvas = new OfficeRasterCanvas(new OfficeRasterImage(1, 1));
        Assert.False(canvas.MeasureFilledContourBounds(new[] { points }, OfficeFillRule.NonZero).IsMeasured);
        using var cancellation = new System.Threading.CancellationTokenSource(); cancellation.Cancel();
        var cancelled = new OfficeRasterCanvas(new OfficeRasterImage(1, 1), font: null, fonts: null, cancellationToken: cancellation.Token);
        Assert.Throws<OperationCanceledException>(() => cancelled.MeasureFilledContourBounds(Array.Empty<List<OfficePoint>>(), OfficeFillRule.NonZero));
    }

    private static List<OfficePoint> Rectangle(double left, double top, double right, double bottom) =>
        new() { new(left, top), new(right, top), new(right, bottom), new(left, bottom) };
}

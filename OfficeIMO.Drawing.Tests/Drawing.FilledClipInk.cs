using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingFilledClipInkTests {
    [Theory]
    [InlineData(OfficeFillRule.EvenOdd, false)]
    [InlineData(OfficeFillRule.NonZero, true)]
    public void ClipHoleUsesDeclaredFillRule(OfficeFillRule rule, bool filled) {
        var outer = Rectangle(0, 0, 20, 20); var inner = Rectangle(5, 5, 15, 15);
        var clip = CreateClip(new[] { outer, inner }, rule);
        var subject = new List<List<OfficePoint>> { Rectangle(7, 7, 13, 13) };
        var canvas = new OfficeRasterCanvas(new OfficeRasterImage(20, 20));
        var bounds = canvas.MeasureFilledContourBounds(subject, OfficeFillRule.NonZero, new[] { clip });
        Assert.True(bounds.IsMeasured); Assert.Equal(filled, bounds.HasInk); Assert.Equal(!filled, bounds.IsClipped);
        var raster = Paint(subject, new[] { outer, inner }, rule);
        Assert.Equal(filled, raster.GetPixel(10, 10).A > 0);
    }

    [Fact]
    public void ConcaveNotchRemovesInkInsideBoundingBox() {
        var polygon = new List<OfficePoint> { new(0, 0), new(20, 0), new(20, 20), new(15, 20), new(15, 5), new(5, 5), new(5, 20), new(0, 20) };
        var subject = new List<List<OfficePoint>> { Rectangle(7, 7, 13, 13) };
        var clip = CreateClip(new[] { polygon }, OfficeFillRule.NonZero);
        var bounds = new OfficeRasterCanvas(new OfficeRasterImage(20, 20)).MeasureFilledContourBounds(subject, OfficeFillRule.NonZero, new[] { clip });
        Assert.True(bounds.IsMeasured && bounds.IsClipped); Assert.False(bounds.HasInk);
        Assert.Equal(0, Paint(subject, new[] { polygon }, OfficeFillRule.NonZero).GetPixel(10, 10).A);
    }

    [Fact]
    public void NestedClipsIntersectAndPreserveDisconnectedIslands() {
        var islands = new[] { Rectangle(0, 0, 5, 20), Rectangle(15, 0, 20, 20) };
        var first = CreateClip(islands, OfficeFillRule.NonZero);
        // Two contours intentionally force fill-aware intersection for the second clip too.
        var second = CreateClip(new[] { Rectangle(0, 10, 20, 20), Rectangle(0, 10, 20, 20) }, OfficeFillRule.NonZero, OfficeTransform.Translate(0, 10));
        var subject = new List<List<OfficePoint>> { Rectangle(0, 0, 20, 20) };
        var bounds = new OfficeRasterCanvas(new OfficeRasterImage(20, 20)).MeasureFilledContourBounds(subject, OfficeFillRule.NonZero, new[] { first, second });
        Assert.True(bounds.IsMeasured && bounds.HasInk && bounds.IsClipped);
        Assert.Equal(0, bounds.Left); Assert.Equal(20, bounds.Right); Assert.Equal(10, bounds.Top); Assert.Equal(20, bounds.Bottom);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SelfIntersectionRetainsFillUnderReflection(bool reflect) {
        var bow = new List<OfficePoint> { new(0, 0), new(20, 20), new(0, 20), new(20, 0) };
        var transform = OfficeTransform.Scale(reflect ? -1 : 1, 1).Then(OfficeTransform.Translate(reflect ? 20 : 0, 0));
        var clip = CreateClip(new[] { bow }, OfficeFillRule.EvenOdd, transform);
        var subject = new List<List<OfficePoint>> { Rectangle(0, 0, 20, 20) };
        var bounds = new OfficeRasterCanvas(new OfficeRasterImage(20, 20)).MeasureFilledContourBounds(subject, OfficeFillRule.NonZero, new[] { clip });
        Assert.True(bounds.IsMeasured && bounds.HasInk && bounds.IsClipped);
        Assert.Equal(0, bounds.Left); Assert.Equal(20, bounds.Right);
        var raster = Paint(subject, new[] { bow }, OfficeFillRule.EvenOdd);
        Assert.Equal(255, raster.GetPixel(10, 2).A); Assert.Equal(0, raster.GetPixel(2, 10).A);
    }

    [Fact]
    public void RepeatedEvenOddBoundaryCancelsAndEmptyClipRemovesInk() {
        var square = Rectangle(0, 0, 20, 20);
        var repeated = square.Concat(square).ToList();
        var canvas = new OfficeRasterCanvas(new OfficeRasterImage(20, 20));
        var clip = CreateClip(new[] { repeated }, OfficeFillRule.EvenOdd);
        var result = canvas.MeasureFilledContourBounds(new[] { square }, OfficeFillRule.NonZero, new[] { clip });
        Assert.True(result.IsMeasured && result.IsClipped); Assert.False(result.HasInk);
        Assert.True(OfficeTextInkClip.TryCreatePath(OfficeClipPath.Empty(), OfficeTransform.Identity, default, out var empty));
        result = canvas.MeasureFilledContourBounds(new[] { square }, OfficeFillRule.NonZero, new[] { empty });
        Assert.True(result.IsMeasured && result.IsClipped); Assert.False(result.HasInk);
    }

    [Fact]
    public void OverBudgetAndUnclosedPathsRemainUnmeasured() {
        var points = Enumerable.Range(0, 600).Select(i => new OfficePoint(10 + 10 * Math.Cos(i * Math.PI / 300), 10 + 10 * Math.Sin(i * Math.PI / 300))).ToArray();
        var commands = new[] { OfficePathCommand.MoveTo(points[0]) }.Concat(points.Skip(1).Select(OfficePathCommand.LineTo)).Concat(new[] { OfficePathCommand.Close() });
        Assert.False(OfficeTextInkClip.TryCreatePath(OfficeClipPath.Path(commands), OfficeTransform.Identity, default, out _));
        var open = OfficeClipPath.Path(OfficePathCommand.MoveTo(new(0, 0)), OfficePathCommand.LineTo(new(20, 0)), OfficePathCommand.LineTo(new(0, 20)));
        Assert.False(OfficeTextInkClip.TryCreatePath(open, OfficeTransform.Identity, default, out _));
    }

    private static OfficeTextInkClip CreateClip(IEnumerable<List<OfficePoint>> contours, OfficeFillRule rule, OfficeTransform? transform = null) {
        var commands = contours.SelectMany(c => new[] { OfficePathCommand.MoveTo(c[0]) }.Concat(c.Skip(1).Select(OfficePathCommand.LineTo)).Concat(new[] { OfficePathCommand.Close() }));
        Assert.True(OfficeTextInkClip.TryCreatePath(OfficeClipPath.Path(commands, rule), transform ?? OfficeTransform.Identity, default, out var clip));
        Assert.NotNull(clip.FilledContours);
        return clip;
    }
    private static OfficeRasterImage Paint(List<List<OfficePoint>> subject, IReadOnlyList<IReadOnlyList<OfficePoint>> clips, OfficeFillRule rule) {
        var image = new OfficeRasterImage(20, 20); var canvas = new OfficeRasterCanvas(image);
        using var scope = rule == OfficeFillRule.EvenOdd ? canvas.PushClipPolygonsEvenOdd(clips) : canvas.PushClipPolygonsNonZero(clips);
        canvas.FillContourPaint(subject, OfficeFillRule.NonZero, (_, _) => OfficeColor.Black);
        return image;
    }
    private static List<OfficePoint> Rectangle(double l, double t, double r, double b) => new() { new(l, t), new(r, t), new(r, b), new(l, b) };
}

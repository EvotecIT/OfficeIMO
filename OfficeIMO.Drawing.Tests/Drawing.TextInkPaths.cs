using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingTextInkPathTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ConvexClipUsesItsEdgesUnderRotationAndReflection(bool reflected) {
        var path = Polygon(new(0, 0), new(10, 0), new(0, 10));
        var transform = OfficeTransform.Scale(reflected ? -1 : 1, 1)
            .Then(OfficeTransform.RotateDegrees(25)).Then(OfficeTransform.Translate(30, 40));
        Assert.True(OfficeTextInkClip.TryCreateConvexPath(path, transform, default, out var clip));
        var outside = new[] { new OfficePoint(6, 6), new OfficePoint(8, 6), new OfficePoint(8, 8), new OfficePoint(6, 8) }
            .Select(transform.TransformPoint).ToList();
        bool cropped = false; long work = 1000;
        Assert.Empty(clip.Apply(outside, ref cropped, ref work, default));
        Assert.True(cropped);
    }

    [Fact]
    public void RoundedClipRemovesInkInsideItsBoxButOutsideItsCurve() {
        Assert.True(OfficeTextInkClip.TryCreateConvexPath(OfficeClipPath.RoundedRectangle(20, 20, 10),
            OfficeTransform.Identity, default, out var clip));
        var corner = new List<OfficePoint> { new(0, 0), new(1, 0), new(1, 1), new(0, 1) };
        bool cropped = false; long work = 10000;
        Assert.Empty(clip.Apply(corner, ref cropped, ref work, default));
        Assert.True(cropped);
    }

    [Fact]
    public void ConcaveIntersectingAndMultiContourPathsRemainUnqualified() {
        var concave = Polygon(new(0, 0), new(10, 0), new(5, 5), new(10, 10), new(0, 10));
        var star = Polygon(new(0, -10), new(6, 8), new(-10, -3), new(10, -3), new(-6, 8));
        var multiple = OfficeClipPath.Path(Polygon(new(0, 0), new(10, 0), new(0, 10)).Commands
            .Concat(Polygon(new(0, 0), new(5, 0), new(0, 5)).Commands));
        var repeated = Polygon(new(0, 0), new(10, 0), new(10, 10), new(0, 10),
            new(0, 0), new(10, 0), new(10, 10), new(0, 10));
        foreach (var path in new[] { concave, star, multiple, repeated })
            Assert.False(OfficeTextInkClip.TryCreateConvexPath(path, OfficeTransform.Identity, default, out _));
    }

    private static OfficeClipPath Polygon(params OfficePoint[] points) => OfficeClipPath.Path(
        new[] { OfficePathCommand.MoveTo(points[0]) }.Concat(points.Skip(1).Select(OfficePathCommand.LineTo))
            .Concat(new[] { OfficePathCommand.Close() }));
}

using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public class DrawingPathMeasureTests {
    [Fact]
    public void PolylineSlicingSkipsRepeatedPointsAndClampsToTheOriginalRoute() {
        var path = new OfficePathMeasure(new[] { OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(0, 0),
            OfficePathCommand.LineTo(30, 0), OfficePathCommand.LineTo(30, 40) });
        Assert.Equal(70, path.Length); Assert.Equal(new OfficePoint(30, 20), path.PointAtLength(50));
        var sliced = path.Slice(20, 50);
        Assert.Equal(new OfficePoint(20, 0), sliced[0].Point); Assert.Equal(new OfficePoint(30, 20), sliced[^1].Point);
        Assert.Equal(new OfficePoint(0, 0), path.PointAtLength(-10)); Assert.Equal(new OfficePoint(30, 40), path.PointAtLength(100));
        Assert.Empty(path.Slice(50, 20)); Assert.Empty(path.Slice(70, 80));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CurveSlicesRetainAnalyticGeometryAndUseArcLengthInsteadOfParameter(bool cubic) {
        var command = cubic ? OfficePathCommand.CubicBezierTo(0, 0, 0, 0, 100, 0) : OfficePathCommand.QuadraticBezierTo(0, 0, 100, 0);
        var path = new OfficePathMeasure(new[] { OfficePathCommand.MoveTo(0, 0), command });
        var slice = path.Slice(25, 75);
        Assert.InRange(slice[0].Point.X, 24.98, 25.02); Assert.InRange(slice[^1].Point.X, 74.98, 75.02);
        Assert.Equal(command.Kind, slice[^1].Kind);
        Assert.Equal(0, slice[^1].Point.Y); Assert.Equal(0, slice[^1].ControlPoint1.Y);
        var middle = new OfficePathMeasure(slice).PointAtLength(25);
        Assert.InRange(middle.X, 49.97, 50.03);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TrimmingBothEndsKeepsTheOriginalCurvedLocus(bool cubic) {
        var command = cubic ? OfficePathCommand.CubicBezierTo(100D/3, 100, 200D/3, 100, 100, 0) : OfficePathCommand.QuadraticBezierTo(50, 100, 100, 0);
        var path = new OfficePathMeasure(new[] { OfficePathCommand.MoveTo(0, 0), command });
        var slice = path.Slice(path.Length*.2, path.Length*.8);
        OfficePoint start = slice[0].Point; var curve = slice[1];
        // These original curves have x=100t and y=k*t*(1-t), an independent analytic oracle.
        double a = start.X/100, b = curve.Point.X/100;
        var samples = cubic ? OfficeGeometry.CreateCubicBezierPoints(start, curve.ControlPoint1, curve.ControlPoint2, curve.Point, 8) :
            OfficeGeometry.CreateQuadraticBezierPoints(start, curve.ControlPoint1, curve.Point, 8);
        for (int i = 0; i < samples.Count; i++) {
            double t = a + (b - a)*(i + 1)/8;
            Assert.Equal(100*t, samples[i].X, 8);
            Assert.Equal((cubic ? 300 : 200)*t*(1 - t), samples[i].Y, 8);
        }
    }

    [Fact]
    public void MeasurementRejectsUnsupportedContoursAndBoundsAggregateCurveWork() {
        Assert.Throws<NotSupportedException>(() => new OfficePathMeasure(new[] { OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(10, 0), OfficePathCommand.Close() }));
        Assert.Throws<NotSupportedException>(() => new OfficePathMeasure(new[] { OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(10, 0), OfficePathCommand.MoveTo(20, 0), OfficePathCommand.LineTo(30, 0) }));
        var commands = new List<OfficePathCommand> { OfficePathCommand.MoveTo(0, 0) };
        for (int i = 0; i < 200; i++) commands.Add(OfficePathCommand.CubicBezierTo(1e8, 1e8, -1e8, 1e8, i + 1, 0));
        Assert.Throws<NotSupportedException>(() => new OfficePathMeasure(commands));
    }
}

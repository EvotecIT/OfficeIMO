using System;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public class DrawingPathGeometryTests {
    [Fact]
    public void BoundsUseBothCubicExtremaAndExcludeUnpaintedMoveContours() {
        var bounds = OfficePathGeometry.Bounds(new[] { OfficePathCommand.MoveTo(1000, 1000), OfficePathCommand.Close(),
            OfficePathCommand.MoveTo(0, 0), OfficePathCommand.CubicBezierTo(0, 20, 20, -20, 20, 0) });
        Assert.Equal(0, bounds.Left); Assert.Equal(20, bounds.Right);
        Assert.Equal(-10/Math.Sqrt(3), bounds.Top, 8); Assert.Equal(10/Math.Sqrt(3), bounds.Bottom, 8);
    }
    [Fact]
    public void EndpointTangentsSkipZeroSegmentsAndRepeatedCurveControls() {
        Assert.True(OfficePathGeometry.TryOpenEndpoints(new[] { OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(0, 0),
            OfficePathCommand.CubicBezierTo(0, 0, 10, 0, 10, 10), OfficePathCommand.LineTo(10, 10) }, out var start, out var forward, out var end, out var backward));
        Assert.Equal(new OfficePoint(0, 0), start); Assert.Equal(new OfficePoint(10, 0), forward);
        Assert.Equal(new OfficePoint(10, 10), end); Assert.Equal(new OfficePoint(0, -10), backward);
    }
}

using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class OfficeGeometryArcSecurityTests {
    [Fact]
    public void EllipticalArcRejectsSweepsThatRequireExcessiveSegments() {
        double excessiveSweep = (4096D + 1D) * (Math.PI / 2D);

        ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            OfficeGeometry.CreateEllipticalArcCubicBezierCommands(
                new OfficePoint(1D, 0D),
                radiusX: 1D,
                radiusY: 1D,
                startRadians: 0D,
                sweepRadians: excessiveSweep));

        Assert.Equal("sweepRadians", exception.ParamName);
    }

    [Fact]
    public void EllipticalArcStillProducesTheExpectedBoundedSegments() {
        List<OfficePathCommand> commands = OfficeGeometry.CreateEllipticalArcCubicBezierCommands(
            new OfficePoint(10D, 0D),
            radiusX: 10D,
            radiusY: 5D,
            startRadians: 0D,
            sweepRadians: Math.PI);

        Assert.Equal(2, commands.Count);
    }
}

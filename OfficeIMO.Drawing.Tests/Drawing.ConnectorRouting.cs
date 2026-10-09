using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingConnectorRoutingTests {
    [Theory]
    [InlineData(0D)]
    [InlineData(-0.3D)]
    public void NonpositiveSpacingUsesTheSameDefaultForConstrainedDetours(double spacing) {
        OfficePoint[][] Routes(double step) => OfficeGeometry.EnumerateConstrainedOrthogonalConnectorRoutes(
            new OfficePoint(0, 0), new OfficePoint(4, 0),
            OfficeConnectorDirections.Left, OfficeConnectorDirections.Right,
            step, 3, 0, null, null).ToArray();

        OfficePoint[][] expected = Routes(0.15D);
        OfficePoint[][] actual = Routes(spacing);
        Assert.NotEmpty(expected);
        Assert.Equal(expected.Select(route => route.Length), actual.Select(route => route.Length));
        Assert.Equal(expected.SelectMany(route => route), actual.SelectMany(route => route));
    }
}

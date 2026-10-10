using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void ClipPath_DeclaredCanvasKeepsInsetCoordinatesAndLegacyOverloadKeepsContentBounds() {
        var commands = new[] { OfficePathCommand.MoveTo(25, 10), OfficePathCommand.LineTo(75, 10),
            OfficePathCommand.LineTo(25, 40), OfficePathCommand.Close() };
        OfficeClipPath content = OfficeClipPath.Path(commands);
        OfficeClipPath canvas = OfficeClipPath.Path(100, 50, commands, OfficeFillRule.NonZero);
        Assert.Equal(50D, content.Width); Assert.Equal(30D, content.Height);
        Assert.Equal(new OfficePoint(0, 0), content.Commands[0].Point);
        Assert.Equal(100D, canvas.Width); Assert.Equal(50D, canvas.Height);
        Assert.Equal(commands, canvas.Commands);
        Assert.Equal(OfficeFillRule.NonZero, canvas.Clone().FillRule);
        Assert.Equal(commands, canvas.Clone().Commands);
    }

    [Theory]
    [InlineData(0, 50)]
    [InlineData(100, -1)]
    [InlineData(double.NaN, 50)]
    [InlineData(100, double.PositiveInfinity)]
    public void ClipPath_DeclaredCanvasRequiresFinitePositiveDimensions(double width, double height) {
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeClipPath.Path(width, height,
            new[] { OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(10, 10) }));
    }
}

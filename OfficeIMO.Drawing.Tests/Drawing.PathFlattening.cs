using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void PathFlattening_BudgetDoesNotCollapseFiniteCubicLoopsToAnEndpoint() {
        var commands = new List<OfficePathCommand>();
        for (int i = 0; i < 22000; i++) {
            commands.Add(OfficePathCommand.MoveTo(i, 0));
            commands.Add(OfficePathCommand.CubicBezierTo(i, 2, i + 2, 2, i, 0));
        }
        var contours = OfficePathFlattener.FlattenStrokeProjection(commands, 4);
        Assert.Equal(22000, contours.Count);
        Assert.All(contours, contour => Assert.Contains(contour.Points, point => point.Y > 0));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PathFlattening_ImplicitContourAfterCloseStartsAtTheClosedOrigin(bool projection) {
        var commands = new[] { OfficePathCommand.MoveTo(2, 3), OfficePathCommand.LineTo(10, 3),
            OfficePathCommand.LineTo(2, 10), OfficePathCommand.Close(),
            OfficePathCommand.LineTo(12, 3), OfficePathCommand.LineTo(12, 12) };
        var contours = projection ? OfficePathFlattener.FlattenStrokeProjection(commands, 4)
            : OfficePathFlattener.Flatten(commands, 0, 0, 1);
        Assert.Equal(2, contours.Count); Assert.False(contours[1].Closed);
        Assert.Equal(new[] { new OfficePoint(2, 3), new OfficePoint(12, 3), new OfficePoint(12, 12) }, contours[1].Points);
    }

}

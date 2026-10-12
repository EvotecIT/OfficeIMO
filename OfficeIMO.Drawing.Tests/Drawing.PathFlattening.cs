using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
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

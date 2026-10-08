using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public class DrawingPathContourTests {
    [Fact]
    public void SplittingRetainsCurvesEmptyMovesAndImplicitOriginAfterClose() {
        var curve = OfficePathCommand.CubicBezierTo(10, 20, 30, 40, 50, 60);
        var contours = OfficePathContour.Split(new[] { OfficePathCommand.MoveTo(1, 2), curve, OfficePathCommand.Close(),
            OfficePathCommand.LineTo(70, 80), OfficePathCommand.MoveTo(9, 9), OfficePathCommand.MoveTo(10, 10) });
        Assert.Equal(4, contours.Count); Assert.True(contours[0].IsClosed);
        Assert.Equal(curve, contours[0].Commands[1]); Assert.Equal(new OfficePoint(1, 2), contours[1].Commands[0].Point);
        Assert.Single(contours[2].Commands); Assert.Single(contours[3].Commands);
        var open = contours[0].OpenWithClosingEdge(); Assert.Equal(curve, open[1]);
        Assert.Equal(OfficePathCommand.LineTo(1, 2), open[^1]); Assert.Equal(OfficePathCommandKind.Close, contours[0].Commands[^1].Kind);
    }

    [Fact]
    public void OpeningDoesNotRepeatAnExistingClosingEndpoint() {
        var contours = OfficePathContour.Split(new[] { OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(10, 10),
            OfficePathCommand.LineTo(0, 0), OfficePathCommand.Close(), OfficePathCommand.Close() });
        var contour = Assert.Single(contours); Assert.Equal(3, contour.OpenWithClosingEdge().Count);
        Assert.Single(OfficePathContour.Split(new[] { OfficePathCommand.MoveTo(1, 1), OfficePathCommand.Close() })[0].OpenWithClosingEdge());
    }
}

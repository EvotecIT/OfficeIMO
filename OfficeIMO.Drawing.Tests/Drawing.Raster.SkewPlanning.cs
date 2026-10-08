using System;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingRasterSkewPlanningTests {
    [Theory]
    [InlineData(45, 45)]
    [InlineData(-45, -45)]
    [InlineData(15, 75)]
    public void SingularSkewRejectsDuringFramePlanning(double x, double y) {
        var image = new OfficeRasterImage(5, 3);
        var frames = new OfficeRasterFrames(new[] { new OfficeRasterFrame(image) });
        bool mapped = false;

        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterTransforms.GetSkewedSize(image, x, y));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeRasterTransforms.Skew(image, x, y));
        Assert.Throws<ArgumentOutOfRangeException>(() => frames.Transform(source => {
            mapped = true;
            return OfficeRasterTransforms.Skew(source, x, y);
        }, source => OfficeRasterTransforms.GetSkewedSize(source, x, y)));
        Assert.False(mapped);
    }

    [Theory]
    [InlineData(45, -45)]
    [InlineData(20, 10)]
    public void InvertibleSkewUsesThePlannedCanvas(double x, double y) {
        var image = new OfficeRasterImage(5, 3);
        var size = OfficeRasterTransforms.GetSkewedSize(image, x, y);
        var result = OfficeRasterTransforms.Skew(image, x, y);

        Assert.Equal(size.Width, result.Width);
        Assert.Equal(size.Height, result.Height);
    }
}

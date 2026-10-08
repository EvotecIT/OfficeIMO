using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class TextDecorationGeometryTests {
    [Theory]
    [InlineData(OfficeTextDecorationStyle.Single, 3D)]
    [InlineData(OfficeTextDecorationStyle.Double, 9D)]
    [InlineData(OfficeTextDecorationStyle.Dashed, 3D)]
    [InlineData(OfficeTextDecorationStyle.Dotted, 3D)]
    [InlineData(OfficeTextDecorationStyle.Wavy, 9D)]
    public void ExplicitVectorBandsKeepBoundsColorAndIndependentPattern(OfficeTextDecorationStyle style, double height) {
        OfficeShape shape = OfficeTextDecorationGeometry.CreateHorizontalBand(97D, 3D, style, OfficeColor.Red);
        Assert.Equal(97D, shape.Width);
        Assert.Equal(height, shape.Height);
        Assert.Equal(OfficeColor.Red, shape.FillColor);
        Assert.Null(shape.StrokeColor);
        Assert.NotEmpty(shape.PathCommands);
        Assert.All(shape.PathCommands, c => {
            Assert.InRange(c.Point.X, 0D, 97D);
            Assert.InRange(c.Point.Y, 0D, height);
            Assert.InRange(c.ControlPoint1.X, 0D, 97D);
            Assert.InRange(c.ControlPoint1.Y, 0D, height);
            Assert.InRange(c.ControlPoint2.X, 0D, 97D);
            Assert.InRange(c.ControlPoint2.Y, 0D, height);
        });
    }

    [Fact]
    public void DecorationGeometryRejectsUnboundedOrInvalidBands() {
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeTextDecorationGeometry.CreateHorizontalBand(double.PositiveInfinity, 1D, OfficeTextDecorationStyle.Single, OfficeColor.Red));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeTextDecorationGeometry.CreateHorizontalBand(20D, 0D, OfficeTextDecorationStyle.Single, OfficeColor.Red));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeTextDecorationGeometry.CreateHorizontalBand(1e9D, 1D, OfficeTextDecorationStyle.Dotted, OfficeColor.Red));
    }
}

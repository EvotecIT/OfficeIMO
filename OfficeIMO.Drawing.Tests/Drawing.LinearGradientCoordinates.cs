using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Theory]
    [InlineData("pad", 50, 229, 26)]
    [InlineData("pad", 350, 76, 179)]
    [InlineData("repeat", 550, 229, 26)]
    [InlineData("reflect", 550, 26, 229)]
    public void OfficeSvgDrawingReader_PreservesShearedUserSpaceGradientNormals(string spread, int x, int red, int blue) {
        // For this source vector and shear, equal-color lines are vertical in
        // user space. At a fixed X both tested Y positions must have one color.
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='600' height='200'>"
            + "<defs><linearGradient id='paint' gradientUnits='userSpaceOnUse' x1='0' y1='0' x2='400' y2='200' "
            + "gradientTransform='matrix(1 0 .5 1 0 0)' spreadMethod='" + spread + "'>"
            + "<stop stop-color='red'/><stop offset='1' stop-color='blue'/></linearGradient></defs>"
            + "<rect width='600' height='200' fill='url(#paint)'/></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing!);
        foreach (int y in new[] { 30, 170 }) {
            OfficeColor actual = image.GetPixel(x, y);
            Assert.InRange((int)actual.R, red - 1, red + 1);
            Assert.InRange((int)actual.B, blue - 1, blue + 1);
        }
    }

    [Fact]
    public void OfficeDrawingRasterRenderer_ShearTransformsEqualColorLinesWithTheShape() {
        var drawing = new OfficeDrawing(400D, 200D);
        OfficeShape shape = OfficeShape.Rectangle(200D, 100D);
        shape.FillGradient = OfficeLinearGradient.DiagonalDown(OfficeColor.Red, OfficeColor.Blue);
        shape.Transform = new OfficeTransform(1D, 0D, 2D, 1D, 0D, 0D);
        drawing.AddShape(shape, 0D, 0D);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing);
        // Local normalized x+y stays constant under this physical shear at a
        // fixed destination X, independently of the destination Y coordinate.
        OfficeColor top = image.GetPixel(180, 20);
        OfficeColor bottom = image.GetPixel(180, 70);
        Assert.InRange((int)top.R, 139, 141);
        Assert.InRange((int)top.B, 114, 116);
        Assert.Equal(top, bottom);
    }

    [Fact]
    public void OfficeSvgDrawingReader_DegenerateGradientTransformIsDiagnosed() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='20' height='10'><defs>"
            + "<linearGradient id='g' gradientTransform='scale(0 1)'><stop stop-color='red'/><stop offset='1' stop-color='blue'/></linearGradient>"
            + "</defs><rect width='20' height='10' fill='url(#g)'/></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out _, out int unsupported));
        Assert.True(unsupported > 0);
    }

    [Fact]
    public void OfficeSvgDrawingReader_SmallInvertibleGradientTransformRetainsItsColorField() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='20' height='10'><defs>"
            + "<linearGradient id='g' gradientUnits='userSpaceOnUse' x2='2000000000' y2='0' gradientTransform='scale(.00000001)'>"
            + "<stop stop-color='red'/><stop offset='1' stop-color='blue'/></linearGradient></defs><rect width='20' height='10' fill='url(#g)'/></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing!);
        Assert.InRange((int)image.GetPixel(5, 5).B, 69, 71);
        Assert.InRange((int)image.GetPixel(15, 5).B, 197, 199);
    }

    [Theory]
    [InlineData("-1e308", "1e308")]
    [InlineData("0", "1e-308")]
    public void OfficeSvgDrawingReader_UnrepresentableGradientVectorIsDiagnosed(string start, string end) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='20' height='10'><defs>"
            + "<linearGradient id='g' x1='" + start + "' y1='0' x2='" + end + "' y2='0'>"
            + "<stop stop-color='red'/><stop offset='1' stop-color='blue'/></linearGradient></defs><rect width='20' height='10' fill='url(#g)'/></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out _, out int unsupported));
        Assert.True(unsupported > 0);
    }
}

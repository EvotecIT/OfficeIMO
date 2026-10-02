using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingTransformedStrokeQualityTests {
    [Theory]
    [InlineData("M10 20 L30 20",20,19,40,2)]
    [InlineData("M10 10 L10 30",18,10,4,20)]
    public void NonUniformTransformScalesTheStrokeOutline(string path, double x, double y, double width, double height) {
        string svg = $"<svg xmlns='http://www.w3.org/2000/svg' width='80' height='40'><path transform='scale(2 1)' d='{path}' fill='none' stroke='black' stroke-width='2' stroke-linecap='butt'/></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg),out var drawing,out int unsupported));
        Assert.Equal(0,unsupported);
        var expected = new OfficeRasterImage(80,40);
        new OfficeRasterCanvas(expected).FillRectangle(x,y,width,height,OfficeColor.Black);
        Assert.Equal(expected.GetPixels(),OfficeDrawingRasterRenderer.Render(drawing!).GetPixels());
    }
}

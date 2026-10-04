using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingSvgGroupOpacityTests {
    [Theory]
    [InlineData("g")]
    [InlineData("a")]
    [InlineData("svg")]
    public void ContainerOpacityCompositesOverlappingChildrenOnce(string element) {
        string content = "<rect width='20' height='20' fill='red'/><rect width='20' height='20' fill='red'/>";
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='20' height='20'><" + element + (element == "a" ? " href='https://example.test/'" : "") + " opacity='0.5'>" + content + "</" + element + "></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out var drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        var pixel = OfficeDrawingRasterRenderer.Render(drawing!).GetPixel(10, 10);
        Assert.InRange(pixel.A, 126, 129);
    }

    [Fact]
    public void ShapeOpacityCompositesFillAndStrokeOnce() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='20' height='20'><rect x='2' y='2' width='16' height='16' fill='red' stroke='red' stroke-width='8' opacity='0.5'/></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out var drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.InRange(OfficeDrawingRasterRenderer.Render(drawing!).GetPixel(3, 10).A, 126, 129);
    }

    [Fact]
    public void RootOpacityAndReferencedGroupOpacityComposeOncePerLevel() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='20' height='20' opacity='0.5'><defs><g id='paint'><rect width='20' height='20' fill='red'/><rect width='20' height='20' fill='red'/></g></defs><use href='#paint' opacity='0.5'/></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out var drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.InRange(OfficeDrawingRasterRenderer.Render(drawing!).GetPixel(10, 10).A, 63, 65);
    }
}

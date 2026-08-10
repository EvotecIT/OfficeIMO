using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingSvgReaderViewportTests {
    [Fact]
    public void SvgReaderPreservesFractionalIntrinsicViewportDimensions() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='10.25' height='5.125' viewBox='0 0 100 50'><rect width='100' height='50'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        Assert.Equal(10.25D, drawing!.Width, 6);
        Assert.Equal(5.125D, drawing.Height, 6);
    }

    [Fact]
    public void SvgReaderDerivesMissingFractionalViewportDimensionFromViewBox() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' height='5.125' viewBox='0 0 100 50'><rect width='100' height='50'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing));
        Assert.NotNull(drawing);
        Assert.Equal(10.25D, drawing!.Width, 6);
        Assert.Equal(5.125D, drawing.Height, 6);
    }
}

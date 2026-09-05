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

    [Fact]
    public void SvgReaderPreservesFractionalIntrinsicDimensionsWithoutViewBox() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='0.4in' height='0.2in'><rect width='100%' height='100%'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing));
        Assert.NotNull(drawing);
        Assert.Equal(38.4D, drawing!.Width, 6);
        Assert.Equal(19.2D, drawing.Height, 6);
    }

    [Fact]
    public void SvgReaderAppliesNestedViewportViewBoxAlignmentAndClipping() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='100' height='60'>"
            + "<svg x='10' y='5' width='40' height='20' viewBox='0 0 10 10' preserveAspectRatio='xMidYMid meet'>"
            + "<rect width='10' height='10' fill='red'/></svg></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        OfficeDrawingEffectGroup viewport = Assert.IsType<OfficeDrawingEffectGroup>(Assert.Single(drawing!.Elements));
        Assert.Equal(10D, viewport.Transform.OffsetX, 6);
        Assert.Equal(5D, viewport.Transform.OffsetY, 6);
        OfficeDrawingGroup clip = Assert.IsType<OfficeDrawingGroup>(Assert.Single(viewport.Drawing.Elements));
        OfficeDrawingEffectGroup fitted = Assert.IsType<OfficeDrawingEffectGroup>(Assert.Single(clip.Drawing.Elements));
        Assert.Equal(2D, fitted.Transform.M11, 6);
        Assert.Equal(2D, fitted.Transform.M22, 6);
        Assert.Equal(10D, fitted.Transform.OffsetX, 6);
        Assert.Equal(0D, fitted.Transform.OffsetY, 6);
        Assert.Single(fitted.Drawing.Shapes);
    }

    [Fact]
    public void SvgReaderResolvesNestedViewportPercentageGeometryAgainstItsParent() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='120' height='80'>"
            + "<svg x='25%' y='25%' width='50%' height='50%'><circle cx='50%' cy='50%' r='10' fill='blue'/></svg></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        OfficeDrawingEffectGroup viewport = Assert.IsType<OfficeDrawingEffectGroup>(Assert.Single(drawing!.Elements));
        Assert.Equal(30D, viewport.Transform.OffsetX, 6);
        Assert.Equal(20D, viewport.Transform.OffsetY, 6);
        Assert.Equal(60D, viewport.Drawing.Width, 6);
        Assert.Equal(40D, viewport.Drawing.Height, 6);
    }
}

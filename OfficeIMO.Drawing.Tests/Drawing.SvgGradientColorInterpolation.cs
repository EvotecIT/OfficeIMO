using System;
using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class SvgGradientColorInterpolationTests {
    [Theory]
    [InlineData("color-interpolation='linearRGB'", "", false, true)]
    [InlineData("", "style='color-interpolation:linearRGB'", false, true)]
    [InlineData("color-interpolation='linearRGB'", "color-interpolation='sRGB'", false, false)]
    [InlineData("color-interpolation='linearRGB'", "", true, true)]
    public void GradientColorSpaceInheritsAndSurvivesSvgExport(string rootStyle, string gradientStyle, bool path, bool linear) {
        string shape = path ? "<path d='M10,10H110V50H10Z'" : "<rect x='10' y='10' width='100' height='40'";
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='120' height='60' " + rootStyle + "><defs><linearGradient id='g' " + gradientStyle + "><stop offset='0' stop-color='red'/><stop offset='1' stop-color='blue'/></linearGradient></defs>" + shape + " fill='url(#g)'/></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out var drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        var original = OfficeDrawingRasterRenderer.Render(drawing!);
        string exported = OfficeDrawingSvgExporter.ToSvg(drawing!, 1, OfficeSvgSizeUnit.Pixel);
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(exported), out var imported, out unsupported));
        Assert.Equal(0, unsupported);
        foreach (var image in new[] { original, OfficeDrawingRasterRenderer.Render(imported!) }) {
            var color = image.GetPixel(59, 30);
            Assert.InRange(color.R, linear ? 187 : 127, linear ? 190 : 130);
            Assert.InRange(color.B, linear ? 186 : 125, linear ? 189 : 128);
        }
    }

    [Fact]
    public void GradientHrefDoesNotInheritTemplateColorInterpolation() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='100' height='20'><defs><linearGradient id='template' color-interpolation='linearRGB'><stop offset='0' stop-color='red'/><stop offset='1' stop-color='blue'/></linearGradient><linearGradient id='used' href='#template'/></defs><rect width='100' height='20' fill='url(#used)'/></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out var drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        var color = OfficeDrawingRasterRenderer.Render(drawing!).GetPixel(49, 10);
        Assert.InRange(color.R, 127, 130);
        Assert.InRange(color.B, 125, 128);
    }
}

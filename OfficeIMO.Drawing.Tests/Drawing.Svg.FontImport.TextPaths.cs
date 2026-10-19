using System;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class DrawingSvgMixedRangeFontTests {
    [Theory]
    [InlineData("M10 55 H225")]
    [InlineData("M10 60 C60 10 170 10 225 60")]
    public void MixedFontTextPathUsesIndividualAdvancesAndRotationCenters(string path) {
        string css = Face(ManagedTextShapingTestAssets.CreateFontWithAdvance(500, 'A'), "U+0041")
            + Face(ManagedTextShapingTestAssets.CreateFontWithAdvance(1500, 'B'), "U+0042");
        OfficeDrawing Read(string content) {
            byte[] bytes = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='240' height='100'>"
                + "<defs><style>" + css + "</style><path id='p' d='" + path + "'/></defs>"
                + "<text font-family='MixedFixture' font-size='30'><textPath href='#p' startOffset='10'>"
                + content + "</textPath></text></svg>");
            Assert.True(OfficeSvgDrawingReader.TryRead(bytes, out OfficeDrawing? drawing, out int unsupported));
            Assert.Equal(0, unsupported);
            return drawing!;
        }
        var actual = OfficeDrawingRasterRenderer.Render(Read("AB"));
        var control = OfficeDrawingRasterRenderer.Render(Read("A<tspan>B</tspan>"));
        Assert.Contains(actual.GetPixels(), value => value != 0);
        Assert.Equal(control.GetPixels(), actual.GetPixels());
    }
}

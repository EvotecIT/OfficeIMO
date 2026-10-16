using System;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingSvgMixedRangeFontTests {
    [Theory]
    [InlineData("start", "fill='black'")]
    [InlineData("end", "fill='black'")]
    [InlineData("start", "fill='none' stroke='black' stroke-width='1'")]
    [InlineData("end", "fill='none' stroke='black' stroke-width='1'")]
    [InlineData("start", "fill='url(#g)'")]
    [InlineData("end", "fill='url(#g)'")]
    public void EmbeddedMixedRangeTextMatchesCompleteFontMeasurementAndPaint(string anchor, string paint) {
        byte[] first = ManagedTextShapingTestAssets.CreateFont('A');
        byte[] complete = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphOutlines('A', 'B');
        string splitCss = Face(first, "U+0041") + Face(complete, "U+0042");
        byte[] actualBytes = Document(splitCss, anchor, paint);
        byte[] controlBytes = Document(Face(complete, null), anchor, paint);

        Assert.True(OfficeSvgDrawingReader.TryRead(actualBytes, out OfficeDrawing? actual, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.True(OfficeSvgDrawingReader.TryRead(controlBytes, out OfficeDrawing? control, out int controlUnsupported));
        Assert.Equal(0, controlUnsupported);
        byte[] pixels = OfficeDrawingRasterRenderer.Render(actual!).GetPixels();
        Assert.Contains(pixels, value => value != 0);
        Assert.Equal(OfficeDrawingRasterRenderer.Render(control!).GetPixels(), pixels);
    }

    private static string Face(byte[] font, string? range) =>
        "@font-face{font-family:MixedFixture;src:url(data:font/ttf;base64," + Convert.ToBase64String(font)
        + ");" + (range == null ? "" : "unicode-range:" + range + ";") + "}";

    private static byte[] Document(string css, string anchor, string paint) => Encoding.UTF8.GetBytes(
        "<svg xmlns='http://www.w3.org/2000/svg' width='240' height='70'><defs><style>" + css + "</style>"
        + "<linearGradient id='g'><stop offset='0' stop-color='red'/><stop offset='1' stop-color='blue'/></linearGradient></defs>"
        + "<text x='110' y='45' text-anchor='" + anchor + "' font-family='MixedFixture' font-size='30' " + paint
        + ">AB<tspan fill='green'>A</tspan></text></svg>");
}

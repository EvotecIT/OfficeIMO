using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingSvgReaderDefaultFontTests {
    [Fact]
    public void SvgReader_UsesCallerDefaultForTextWithoutAuthoredFontFamily() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 30'>"
            + "<text x='2' y='20'>Browser default</text></svg>";
        var options = new OfficeSvgDrawingReaderOptions { DefaultFontFamily = "serif" };

        Assert.True(OfficeSvgDrawingReader.TryRead(
            Encoding.UTF8.GetBytes(svg), options, out OfficeDrawing? drawing, out int unsupported));

        Assert.Equal(0, unsupported);
        Assert.Equal("serif", Assert.Single(drawing!.Elements.OfType<OfficeDrawingText>()).Font.FamilyName);
    }

    [Fact]
    public void SvgReader_RejectsAnInvalidCallerDefaultFontFamily() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'/>";

        Assert.False(OfficeSvgDrawingReader.TryRead(
            Encoding.UTF8.GetBytes(svg),
            new OfficeSvgDrawingReaderOptions { DefaultFontFamily = " " },
            out _));
    }
}

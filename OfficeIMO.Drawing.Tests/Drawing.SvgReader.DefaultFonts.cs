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

    [Fact]
    public void SvgReader_UsesCallerDefaultInsidePatternDefinitions() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 40 20'>"
            + "<defs><pattern id='p' patternUnits='userSpaceOnUse' width='20' height='20'>"
            + "<text x='1' y='15'>A</text></pattern></defs>"
            + "<rect width='40' height='20' fill='url(#p)'/></svg>";
        var options = new OfficeSvgDrawingReaderOptions { DefaultFontFamily = "serif" };

        Assert.True(OfficeSvgDrawingReader.TryRead(
            Encoding.UTF8.GetBytes(svg), options, out OfficeDrawing? drawing, out int unsupported));

        Assert.Equal(0, unsupported);
        OfficeDrawingEffectGroup elementGroup = Assert.Single(drawing!.Elements.OfType<OfficeDrawingEffectGroup>());
        OfficeDrawingEffectGroup patternHost = Assert.Single(elementGroup.Drawing.Elements.OfType<OfficeDrawingEffectGroup>());
        OfficeDrawingEffectGroup transformedPattern = Assert.Single(patternHost.Drawing.Elements.OfType<OfficeDrawingEffectGroup>());
        OfficeDrawingGroup clippedPattern = Assert.Single(transformedPattern.Drawing.Elements.OfType<OfficeDrawingGroup>());
        OfficeDrawingTilingPattern pattern = Assert.Single(clippedPattern.Drawing.Elements.OfType<OfficeDrawingTilingPattern>());
        Assert.Contains("font-family=\"serif\"", OfficeDrawingSvgExporter.ToSvg(pattern.Tile), StringComparison.Ordinal);
    }
}

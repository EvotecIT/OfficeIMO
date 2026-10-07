using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("", "math", OfficeFontStyle.Regular)]
    [InlineData("font-family:inherit;font-weight:inherit;font-style:inherit", "Surrounding", OfficeFontStyle.Bold | OfficeFontStyle.Italic)]
    [InlineData("font-family:Authored;font-weight:bold;font-style:italic", "Authored", OfficeFontStyle.Bold | OfficeFontStyle.Italic)]
    [InlineData("font-family:initial;font-weight:initial;font-style:initial", "serif", OfficeFontStyle.Regular)]
    [InlineData("all:initial;font-size:40px", "serif", OfficeFontStyle.Regular)]
    [InlineData("font-family:unset;font-weight:unset;font-style:unset", "Surrounding", OfficeFontStyle.Bold | OfficeFontStyle.Italic)]
    [InlineData("font-family:revert;font-weight:revert;font-style:revert", "math", OfficeFontStyle.Regular)]
    public void HtmlMathMl_DefaultFontResetsAmbientStyleAndPreservesExplicitOverrides(string css, string family, OfficeFontStyle style) {
        var options = new HtmlRenderOptions {
            AllowSystemFontFallback = false, Margins = HtmlRenderMargins.All(0D)
        };
        foreach (string name in new[] { "math", "serif", "Surrounding", "Authored" })
            options.Fonts.Add(name, ManagedTextShapingTestAssets.CreateFont('x'));
        string html = "<body style='font:italic bold 40px Surrounding'>"
            + "<math style='" + css + "'><mtext>x</mtext></math></body>";
        HtmlRenderDrawing math = Assert.Single(HtmlRenderTestDriver.Render(html, options).Pages[0].Visuals.OfType<HtmlRenderDrawing>());
        OfficeDrawingText glyph = Assert.Single(math.Drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Equal(family, glyph.Font.FamilyName);
        Assert.Equal(style, glyph.Font.Style);
        Assert.Equal(40D, glyph.Font.Size);
        Assert.True(math.Drawing.Fonts.TryResolveFaceForText("x", family, style, out _));
    }

    [Fact]
    public void HtmlMathMl_GenericMathUsesSuppliedMathematicalFaceWithoutInstalledFonts() {
        var options = new HtmlRenderOptions {
            AllowSystemFontFallback = false, Margins = HtmlRenderMargins.All(0D)
        };
        options.Fonts.Add("STIX Two Math", ManagedTextShapingTestAssets.CreateFont('x'));
        HtmlRenderDrawing math = Assert.Single(HtmlRenderTestDriver.Render("<math><mtext>x</mtext></math>", options)
            .Pages[0].Visuals.OfType<HtmlRenderDrawing>());
        OfficeDrawingText glyph = Assert.Single(math.Drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Equal("math", glyph.Font.FamilyName);
        Assert.True(math.Drawing.Fonts.TryResolveFaceForText("x", "math", OfficeFontStyle.Regular, out OfficeFontFace? face));
        Assert.Equal("STIX Two Math", face!.FamilyName);
    }
}

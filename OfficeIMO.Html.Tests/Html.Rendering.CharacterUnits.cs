using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using System.Globalization;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlCharacterUnitTests {
    [Theory]
    [InlineData("Arial", "3ch", 3D, 0D)]
    [InlineData("Courier New", "3ch", 3D, 0D)]
    [InlineData("Arial", "calc(2ch + 7px)", 2D, 7D)]
    public void CharacterMarginUsesTheRenderedZeroAdvance(string family, string length, double count, double extra) {
        var text = Render("<div style='font-family:" + family + ";font-size:20px'><div>0</div><div style='margin-left:" + length + "'>Body</div></div>");
        var zero = Find(text, "0");
        Assert.Equal(zero.TextAdvanceWidth!.Value * count + extra, Find(text, "Body").X, 3);
    }

    [Fact]
    public void CharacterWidthConstrainsTheActualBox() {
        var text = Render("<div style='font-family:Courier New;font-size:20px'><div>0</div><div style='width:5ch;overflow-wrap:anywhere'>1234567890</div></div>");
        var zero = Find(text, "0");
        var digits = text.Where(item => item.Text != "0").ToArray();
        Assert.Equal(2, digits.Select(item => item.Y).Distinct().Count());
        Assert.All(digits, item => Assert.True(item.TextAdvanceWidth <= zero.TextAdvanceWidth * 5D + 0.01D));
    }

    [Fact]
    public void FontSizeCharacterUnitUsesParentFontMetrics() {
        var text = Render("<div style='font-family:Courier New;font-size:20px'><div>0</div><div style='font-size:2ch;font-family:Arial'>Child</div></div>");
        Assert.Equal(Find(text, "0").TextAdvanceWidth!.Value * 2D, Find(text, "Child").Font.Size, 3);
    }

    [Theory]
    [InlineData("<div style='font-size:0;width:10ch'>Hidden</div><p>Visible</p>")]
    [InlineData("<style>:root{--measure:20ch}</style><div style='font-size:0'>Hidden<span style='font-size:16px'>Visible</span></div>")]
    [InlineData("<div style='font-size:0'><span style='font-size:2ch'>Hidden</span><span style='font-size:16px'>Visible</span></div>")]
    public void ZeroSizeCharacterMeasurementPreservesVisiblePdfDescendants(string html) {
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions {
            ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        });

        string text = OfficeIMO.Pdf.PdfReadDocument.Open(pdf).ExtractText();
        Assert.Equal("Visible", text.Trim());
    }

    [Fact]
    public void ZeroSizeCharacterEdgesAndInheritedIndentDoNotShiftVisibleDescendants() {
        HtmlRenderText visible = Assert.Single(Render(
            "<div style='font-size:0;margin-left:3ch;padding-left:2ch;text-indent:1ch'>Hidden"
            + "<span style='font-size:16px'>Visible</span></div>"));

        Assert.Equal("Visible", visible.Text);
        Assert.Equal(16D, visible.Font.Size, 3);
        Assert.Equal(0D, visible.X, 3);
    }

    [Fact]
    public void InheritedCharacterIndentKeepsDeclaringFontMetrics() {
        var text = Render("<div style='font-family:Courier New;font-size:20px;text-indent:2ch'><div style='text-indent:0'>0</div><div style='font-family:Arial;font-size:10px'>Child</div></div>");
        Assert.Equal(Find(text, "0").TextAdvanceWidth!.Value * 2D, Find(text, "Child").X, 3);
    }

    [Fact]
    public void DeferredPositioningUsesTheSameCharacterAdvance() {
        var text = Render("<div style='font-family:Courier New;font-size:20px'><div>0</div><div style='position:relative;left:3ch'>Shifted</div></div>");
        Assert.Equal(Find(text, "0").TextAdvanceWidth!.Value * 3D, Find(text, "Shifted").X, 3);
    }

    [Fact]
    public void ScopedFontCharacterAdvanceExcludesLetterSpacing() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('0', 'B');
        string css = "<style>@font-face{font-family:Scoped;src:url(data:font/ttf;base64," + Convert.ToBase64String(font) + ")}</style>";
        var text = Render(css + "<div style='font-family:Scoped;font-size:40px'><div>0</div><div style='letter-spacing:9px;margin-left:3ch'>B</div></div>");
        Assert.Equal("Scoped", Find(text, "0").Font.FamilyName);
        Assert.Equal(Find(text, "0").TextAdvanceWidth!.Value * 3D, Find(text, "B").X, 3);
    }

    [Fact]
    public void CharacterAdvanceUsesTheActiveOpenTypeFeature() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithMultipleSubstitution('0', featureTag: "tnum");
        string css = "<style>@font-face{font-family:Scoped;src:url(data:font/ttf;base64," + Convert.ToBase64String(font) + ")}</style>";
        var text = Render(css + "<div style='font-family:Scoped;font-size:40px;font-variant-numeric:tabular-nums'><div>0</div><div style='margin-left:3ch'>0</div></div>");
        Assert.Equal(2, text.Length);
        Assert.Equal(text[0].TextAdvanceWidth!.Value * 3D, text[1].X, 3);
    }

    [Theory]
    [InlineData("border:1ch solid red;border-radius:2ch")]
    [InlineData("box-shadow:1ch 2ch 0 red")]
    [InlineData("text-shadow:1ch 2ch 0 red")]
    [InlineData("transform:translate(2ch,1ch);transform-origin:1ch 2ch")]
    [InlineData("clip-path:inset(1ch 2ch)")]
    [InlineData("background:linear-gradient(to right,red 1ch,blue 2ch)")]
    [InlineData("background:radial-gradient(circle 2ch at 1ch 2ch,red,blue)")]
    public void CharacterPaintLengthsMatchEquivalentMeasuredPixels(string declarations) {
        double advance = Find(Render("<div style='font-family:Courier New;font-size:20px'>0</div>"), "0").TextAdvanceWidth!.Value;
        string pixels = declarations.Replace("2ch", (2D * advance).ToString("R", CultureInfo.InvariantCulture) + "px")
            .Replace("1ch", advance.ToString("R", CultureInfo.InvariantCulture) + "px");
        string Draw(string style) {
            var rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
                "<div style='font-family:Courier New;font-size:20px;width:140px;height:80px;" + style + "'>Body</div>"),
                new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 300D, Margins = HtmlRenderMargins.All(0D) });
            return OfficeDrawingSvgExporter.ToSvg(rendered.Pages[0].CreateDrawing());
        }
        Assert.Equal(Draw(pixels), Draw(declarations));
    }

    [Theory]
    [InlineData("letter-spacing", "1ch", "0 0")]
    [InlineData("word-spacing", "1ch", "0 0")]
    [InlineData("line-height", "3ch", "0<br/>0")]
    [InlineData("text-shadow", "1ch 1ch red", "0")]
    [InlineData("text-shadow", "1ch 1ch currentColor", "0")]
    [InlineData("text-shadow", "1ch 1ch", "0")]
    [InlineData("line-height", "2", "0<br/>0")]
    [InlineData("line-height", "normal", "0<br/>0")]
    [InlineData("border-spacing", "1ch", "<table><tr><td>0</td><td>0</td></tr></table>")]
    public void InheritedCharacterLengthsKeepTheDeclaringFont(string property, string value, string content) {
        double advance = Find(Render("<div style='font-family:Courier New;font-size:20px'>0</div>"), "0").TextAdvanceWidth!.Value;
        string resolved = value.Replace("3ch", (3D * advance).ToString("R", CultureInfo.InvariantCulture) + "px")
            .Replace("1ch", advance.ToString("R", CultureInfo.InvariantCulture) + "px");
        string Draw(string childDeclaration) {
            var rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
                "<div style='font-family:Courier New;font-size:20px;" + property + ":" + value
                + "'><div style='font-family:Arial;font-size:10px;color:blue;" + childDeclaration + "'>" + content + "</div></div>"),
                new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 300D, Margins = HtmlRenderMargins.All(0D) });
            return OfficeDrawingSvgExporter.ToSvg(rendered.Pages[0].CreateDrawing());
        }
        Assert.Equal(Draw(property + ":" + resolved), Draw(""));
    }

    private static HtmlRenderText Find(HtmlRenderText[] text, string value) => Assert.Single(text, item => item.Text == value);
    private static HtmlRenderText[] Render(string html) => HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html),
        new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 300D, Margins = HtmlRenderMargins.All(0D) })
        .Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().ToArray();
}

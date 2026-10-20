using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("AA END-12345", false, false, "")]
    [InlineData("AA END<span>-12345</span>", false, false, "")]
    [InlineData("AA END-<span>12345</span>", false, false, "")]
    [InlineData("AA END<span>-12345</span>", true, false, "")]
    [InlineData("AA END<span>-12345</span>", false, true, "")]
    [InlineData("AA END-12345", false, false, "word-break:keep-all;")]
    public void HtmlNumericHyphen_UsesTheCurrentLineRemainderAcrossPaintStyles(
        string content, bool floated, bool firstLine, string extraStyle) {
        HtmlRenderDocument rendered = RenderNumericHyphen(content, floated ? 120D : 100D, extraStyle,
            floated, firstLine);

        Assert.Equal(new[] { "AA END-", "12345" }, NumericHyphenLines(rendered));
        Assert.Equal("AA END-12345", NumericHyphenLogicalText(rendered));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("AA -12345", "-12345")]
    [InlineData("AA <span>-12345</span>", "-12345")]
    [InlineData("AA #-12345", "#-12345")]
    public void HtmlNumericHyphen_KeepsLeadingNegativeNumbersTogether(string content, string number) {
        HtmlRenderDocument rendered = RenderNumericHyphen(content, 80D);

        Assert.Equal(new[] { "AA", number }, NumericHyphenLines(rendered));
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlNumericHyphen_NoWrapKeepsTheIdentifierWhole() {
        HtmlRenderDocument rendered = RenderNumericHyphen("AA END-12345", 80D, "white-space:nowrap;");

        Assert.Equal(new[] { "AA END-12345" }, NumericHyphenLines(rendered));
    }

    [Theory]
    [InlineData("END-12345")]
    [InlineData("END<span>-12345</span>")]
    public void HtmlNumericHyphen_MinContentUsesTheSameTailoredBoundary(string content) {
        HtmlRenderDocument rendered = RenderNumericHyphen(content, 200D, "display:inline-block;width:min-content;");
        HtmlRenderShape box = Assert.Single(rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderShape>(), item => item.Source == "div#numeric");

        Assert.Equal(50D, box.Width, 6);
        Assert.Equal(new[] { "END-", "12345" }, NumericHyphenLines(rendered));
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlNumericHyphen_DoesNotSplitTheSurrogateBeforeTheHyphen() {
        HtmlRenderDocument rendered = RenderNumericHyphen("AA \U00010400-12345", 60D);

        Assert.Equal(new[] { "AA \U00010400-", "12345" }, NumericHyphenLines(rendered));
        Assert.Equal("AA \U00010400-12345", NumericHyphenLogicalText(rendered));
        rendered.RequireNoLoss();
    }

    private static HtmlRenderDocument RenderNumericHyphen(string content, double width, string extraStyle = "",
        bool floated = false, bool firstLine = false) {
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous, ViewportWidth = 300D, ViewportHeight = 300D,
            Margins = HtmlRenderMargins.All(0D), AllowSystemFontFallback = false,
            DefaultFontFamily = ManagedTextShapingTestAssets.FamilyName
        };
        var advances = new Dictionary<int, ushort>();
        foreach (char character in " AEND-12345#") advances[character] = 500;
        advances[0x10400] = 500;
        options.Fonts.Add(ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFontWithScalarAdvances(advances));
        string css = "<style>html,body{margin:0;padding:0}#numeric{margin:0;padding:0;border:0;background:white;"
            + "font:20px/24px '" + ManagedTextShapingTestAssets.FamilyName + "';font-kerning:none;"
            + "width:" + width.ToString("R", System.Globalization.CultureInfo.InvariantCulture) + "px;" + extraStyle + "}"
            + (firstLine ? "#numeric::first-line{color:red}" : "") + "</style>";
        string obstruction = floated ? "<span style='float:left;width:20px;height:200px'></span>" : "";
        return HtmlRenderEngine.Render(HtmlConversionDocument.Parse(css + "<div id='numeric'>" + obstruction + content + "</div>"), options);
    }

    // These fixtures have no authored hard breaks. Keep their authored spaces
    // while omitting the scene text projection's soft layout-line separators.
    private static string NumericHyphenLogicalText(HtmlRenderDocument document) =>
        document.Text.Replace("\r", string.Empty).Replace("\n", string.Empty).Trim();

    private static string[] NumericHyphenLines(HtmlRenderDocument document) => document.Pages
        .SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderText>()
        .GroupBy(text => text.Y).OrderBy(line => line.Key)
        .Select(line => string.Concat(line.OrderBy(text => text.X).Select(text => text.Text))).ToArray();
}

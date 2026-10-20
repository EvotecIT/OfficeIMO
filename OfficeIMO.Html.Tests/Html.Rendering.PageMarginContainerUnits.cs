using AngleSharp.Html.Parser;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("font:16px Pinned;line-height:10cqw", "10cqw")]
    [InlineData("font:italic bold 16px/10cqw Pinned", "10cqw")]
    [InlineData("font:16px/10cqw Pinned;line-height:22px", "22px")]
    [InlineData("line-height:10cqw;font:16px Pinned", "normal")]
    [InlineData("line-height:10cqw!important;font:16px/22px Pinned", "10cqw")]
    [InlineData("font:16px/10cqw Pinned!important;line-height:22px", "10cqw")]
    [InlineData("font:16px/10cqw Pinned!important;line-height:revert-layer", "10cqw")]
    [InlineData("font:16px/10cqw Pinned;line-height:-1cqw!important", "10cqw")]
    [InlineData("font:16px/10cqw Pinned;line-height:calc(1cqw +)", "10cqw")]
    [InlineData("font:16px/10cqw Pinned;font:italic 12px/2cqw!important", "10cqw")]
    [InlineData("font:16px/10cqw Pinned;font:12px/-2cqw Other!important", "10cqw")]
    public void HtmlPageMarginContainerUnits_PreserveValidDeclarationsAndTheirCascade(
        string declarations, string expected) {
        var options = PageMarginTypographyOptions();
        using var document = new HtmlParser().ParseDocument(PageMarginTypographySource(declarations));
        HtmlCssPageRuleSet rules = HtmlCssPageSettingsResolver.Apply(document, options, new HtmlDiagnosticReport());
        HtmlCssPageGeometry geometry = rules.ResolveGeometry(1, null, options);

        HtmlCssPageMarginTemplate margin = Assert.Single(rules.ResolveMarginBoxes(1, null, geometry, options)).Value;

        Assert.Equal(expected, margin.LineHeight);
        Assert.Equal(16D, margin.Font.Size, 6);
    }

    [Theory]
    [InlineData("10cqw", 36D)]
    [InlineData("10cqh", 48D)]
    [InlineData("10cqi", 36D)]
    [InlineData("10cqb", 48D)]
    [InlineData("10cqmin", 36D)]
    [InlineData("10cqmax", 48D)]
    [InlineData("calc(10cqw + 4px)", 40D)]
    [InlineData("10vw", 36D)]
    public void HtmlPageMarginContainerUnits_UseThePrintPaperInsteadOfTheLastBodyContainer(
        string leading, double expected) {
        var options = PageMarginTypographyOptions();
        options.PageSize = new OfficePageSize(360D / 96D, 480D / 96D);
        string html = PageMarginTypographySource("font:16px/" + leading + " Pinned", "Hgyp")
            + "<div style='container-type:size;width:100px;height:50px;font-size:32px'>Body container</div>";

        HtmlRenderPage page = Assert.Single(HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options).Pages);

        Assert.Equal(300D, page.Width, 6);
        Assert.Equal(expected, PageMarginTypographyText(page).LineHeight, 6);
    }

    [Fact]
    public void HtmlPageMarginContainerUnits_ShorthandPreservesSizeFaceAndLeadingTogether() {
        var options = PageMarginTypographyOptions();
        options.PageSize = new OfficePageSize(600D / 96D, 480D / 96D);
        string html = PageMarginTypographySource("font:italic bold 2cqw/4cqw Pinned", "Hgyp");

        HtmlRenderPage page = Assert.Single(HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options).Pages);
        HtmlRenderText margin = PageMarginTypographyText(page);

        Assert.Equal(12D, margin.Font.Size, 6);
        Assert.Equal(24D, margin.LineHeight, 6);
        Assert.Equal(OfficeFontStyle.Bold | OfficeFontStyle.Italic, margin.Font.Style);
    }

    [Fact]
    public void HtmlPageMarginContainerUnits_RepeatedBlocksAndLayerRevertRetainTheWinningLeading() {
        const string html = """
            <style>
              @layer base,theme;
              @layer base { @page { size:300px 200px;margin:50px;
                @bottom-center { content:'Hgyp';font:16px/10cqw Pinned; }
                @bottom-center { line-height:22px; }
              } }
              @layer theme { @page { @bottom-center { font:16px/8px Other;font:revert-layer; } } }
              body,p { margin:0;font:16px/20px Pinned; }
            </style><p>BodyOnly</p>
            """;
        HtmlRenderPage page = Assert.Single(HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(html), PageMarginTypographyOptions()).Pages);

        Assert.Equal(22D, PageMarginTypographyText(page).LineHeight, 6);
        Assert.Equal("Pinned", PageMarginTypographyText(page).Font.FamilyName);
    }
}

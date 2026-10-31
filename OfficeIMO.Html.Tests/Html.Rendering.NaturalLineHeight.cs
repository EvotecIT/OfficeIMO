using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("13.3333333333333px", "", "Hgyp<br>Hgyp", 34D)]
    [InlineData("16px", "", "Hgyp<br>Hgyp", 38D)]
    [InlineData("10.6666666666667px", "", "Hgyp<br>Hgyp", 26D)]
    [InlineData("13.3333333333333px", "normal", "Hgyp<br>Hgyp", 34D)]
    [InlineData("13.3333333333333px", "1.2", "Hgyp<br>Hgyp", 31.96875D)]
    [InlineData("13.3333333333333px", "20px", "Hgyp<br>Hgyp", 40D)]
    [InlineData("16px", "", "Hgyp <strong>Hgyp</strong><br>Hgyp", 38D)]
    [InlineData("13.3333333333333px", "", "<strong>Hgyp</strong><br><span style='font-size:10.6666666666667px'>Hgyp</span>", 34D)]
    [InlineData("16px", "", "<span style='font-size:24px'></span>Hgyp<br>Hgyp", 48D)]
    public void HtmlNaturalLineHeight_SelectedFaceAndEmptyStrutsSetFlowHeight(
        string fontSize, string leading, string content, double expectedHeight) {
        // Edge uses these exact licensed face bytes. Normal rounds the face's
        // ascent and descent; authored leading keeps its existing CSS precision.
        string html = NaturalLineSource("<div id='subject' style='font-size:" + fontSize
            + ";line-height:" + leading + "'>" + content + "</div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, NaturalLineOptions());

        Assert.Equal(expectedHeight, NaturalLineShape(rendered, "div#subject").Height, 6);
        Assert.Equal(expectedHeight, NaturalLineShape(rendered, "div#following").Y, 6);
    }

    [Theory]
    [InlineData(false, 19D)]
    [InlineData(true, 19.1875D)]
    public void HtmlNaturalLineHeight_CloningRetainsAnExplicitDefaultIncludingOnePointTwo(
        bool explicitDefault, double expectedHeight) {
        HtmlRenderOptions options = NaturalLineOptions();
        if (explicitDefault) options.DefaultLineHeight = 1.2D;
        string html = NaturalLineSource("<div id='subject'>Hgyp</div>");
        foreach (HtmlRenderOptions copy in new[] { options, options.Clone() }) {
            HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, copy);
            Assert.Equal(expectedHeight, NaturalLineShape(rendered, "div#subject").Height, 6);
        }
    }

    [Theory]
    [InlineData("document")]
    [InlineData("named")]
    [InlineData("standard")]
    public void HtmlNaturalLineHeight_DocumentAndPdfFallbackFacesIncludeTheFontLineGap(string source) {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(800, -200, 100, 200, 'A', 'B');
        var options = new HtmlToPdfOptions(NaturalLineOptions()) {
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        };
        if (source == "named") options.PdfOptions.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("GapFace", font));
        else if (source == "standard") options.FontFamily = new PdfEmbeddedFontFamily("GapFace", font);
        else options.Fonts.Add("GapFace", font);
        string html = NaturalLineSource("<div id='subject' style='font:20px GapFace'>A<br>A</div>");
        HtmlRenderDocument rendered = HtmlConversionDocument.Parse(html).RenderToPdfResult(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options)).RenderResult.Document;

        Assert.Equal(44D, NaturalLineShape(rendered, "div#subject").Height, 6);
        Assert.Equal(44D, NaturalLineShape(rendered, "div#following").Y, 6);
    }

    [Fact]
    public void HtmlNaturalLineHeight_CellsUseTheirParticipatingFontBoxes() {
        string html = NaturalLineSource("<table style='width:600px;border-collapse:collapse'><tbody>"
            + string.Concat(Enumerable.Range(0, 3).Select(index => "<tr><td style='padding:5px 6px;border-bottom:1px solid #aaa'>"
                + "<strong>Hgyp sample " + index + "</strong><br><span style='font-size:10.6666666666667px'>Hgyp subline</span>"
                + "</td><td style='padding:5px 6px;border-bottom:1px solid #aaa'>123</td></tr>"))
            + "</tbody></table>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, NaturalLineOptions());

        Assert.Equal(147D, NaturalLineShape(rendered, "div#following").Y, 6);
        Assert.Equal(3, rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>().Count(text => text.Text.StartsWith("Hgyp sample", StringComparison.Ordinal)));
    }

    [Fact]
    public void HtmlNaturalLineHeight_TableFragmentationUsesActualLineEnds() {
        const string html = "<!doctype html><style>@page{size:200px 42px;margin:0}"
            + "body{margin:0;font:16px Pinned}table{border-collapse:collapse;width:180px}"
            + "td{padding:0;border:0;orphans:1;widows:1}</style>"
            + "<table><tr><td>FIRST<br>SECOND<br>THIRD</td></tr></table>";
        HtmlRenderOptions options = NaturalLineOptions();
        options.HonorCssPageRules = true;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Equal(new[] { "FIRST", "SECOND" }, rendered.Pages[0].Visuals.OfType<HtmlRenderText>().Select(text => text.Text));
        Assert.Equal("THIRD", Assert.Single(rendered.Pages[1].Visuals.OfType<HtmlRenderText>()).Text);
    }

    [Fact]
    public void HtmlNaturalLineHeight_CffFacesSupplyTheirOwnNormalAdvance() {
        HtmlRenderOptions options = NaturalLineOptions();
        options.Fonts.Add("CffFace", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "SourceSansPro-Regular.otf")));
        string html = NaturalLineSource("<div id='subject' style='font-family:CffFace'>Hgyp<br>Hgyp</div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        Assert.Equal(40D, NaturalLineShape(rendered, "div#subject").Height, 6);
        Assert.Equal(40D, NaturalLineShape(rendered, "div#following").Y, 6);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlNaturalLineHeight_StyledGlyphFallbackCombinesEachFacesOwnLeading(bool documentFonts) {
        byte[] regular = ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(700, -300, 100, 300, 'A', 'B');
        byte[] bold = ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(800, -200, 0, 200, 'A', 'C');
        var options = new HtmlToPdfOptions { Margins = HtmlRenderMargins.All(0D),
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
        if (documentFonts) {
            options.Fonts.Add("Scoped", regular);
            options.Fonts.Add("Scoped", bold, OfficeFontStyle.Bold);
        } else options.PdfOptions.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Scoped", regular, bold));
        string html = NaturalLineSource("<div id='subject' style='font:20px Scoped;font-weight:bold'>AB</div>");
        HtmlRenderDocument rendered = HtmlConversionDocument.Parse(html).RenderToPdfResult(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options)).RenderResult.Document;

        // Bold A contributes A16/D4; regular B contributes A15/D7 after
        // its own 2px gap. Their shared baseline is16 and total height23.
        Assert.Equal(23D, NaturalLineShape(rendered, "div#subject").Height, 6);
        Assert.Equal(23D, NaturalLineShape(rendered, "div#following").Y, 6);
    }

    private static HtmlRenderOptions NaturalLineOptions() {
        var options = new HtmlRenderOptions {
            ViewportWidth = 800D, ViewportHeight = 900D,
            PageSize = new OfficePageSize(global::OfficeIMO.OfficeLength.Inches(800D / 96D), global::OfficeIMO.OfficeLength.Inches(900D / 96D)),
            Margins = HtmlRenderMargins.All(0D), HonorCssPageRules = false,
            UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser, AllowSystemFontFallback = false
        };
        string fonts = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Fonts");
        options.Fonts.Add("Pinned", File.ReadAllBytes(Path.Combine(fonts, "OfficeIMOBaselineSans-Regular.ttf")));
        options.Fonts.Add("Pinned", File.ReadAllBytes(Path.Combine(fonts, "OfficeIMOBaselineSans-Bold.ttf")), OfficeFontStyle.Bold);
        return options;
    }

    private static string NaturalLineSource(string content) =>
        "<!doctype html><style>body{margin:0;font:16px Pinned}#subject{background:#eee}"
        + "#following{height:2px;background:red}table,td{margin:0;border-spacing:0}</style>"
        + content + "<div id='following'></div>";

    private static HtmlRenderShape NaturalLineShape(HtmlRenderDocument rendered, string source) =>
        Assert.Single(rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderShape>(),
            shape => shape.Source == source && shape.Shape.FillColor.HasValue);
}

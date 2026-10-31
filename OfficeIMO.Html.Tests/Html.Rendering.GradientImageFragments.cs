using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("linear")]
    [InlineData("radial")]
    [InlineData("conic")]
    [InlineData("svg")]
    public void HtmlGradientImageFragments_BackgroundPatternsDoNotMakeTallBlocksAtomic(string family) {
        string html = GradientImageFragmentSource(
            "<p>PHASE-LEAD</p><section id='fragment-paint'><p>PHASE-START</p></section><p>PHASE-AFTER</p>",
            "#fragment-paint{box-sizing:content-box;width:160px;height:140px;padding:4px;border:2px solid green;"
            + GradientImageFragmentPaint(family) + "}");
        HtmlRenderDocument paged = HtmlRenderTestDriver.Render(html, GradientImageFragmentOptions(HtmlRenderMode.Paged));
        HtmlRenderDocument continuous = HtmlRenderTestDriver.Render(html, GradientImageFragmentOptions(HtmlRenderMode.Continuous));

        Assert.Equal(2, paged.Pages.Count);
        AssertGradientImageFragmentMarkersOnce(paged, "PHASE-LEAD", "PHASE-START", "PHASE-AFTER");
        Assert.DoesNotContain(paged.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment);

        // The second sheet must continue the source tile phase rather than
        // restarting the background at its fragment's physical page origin.
        OfficeRasterImage control = OfficeDrawingRasterRenderer.Render(Assert.Single(continuous.Pages).CreateDrawing());
        for (int pageIndex = 0; pageIndex < paged.Pages.Count; pageIndex++) {
            OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(paged.Pages[pageIndex].CreateDrawing());
            int[] rows = pageIndex == 0 ? new[] { 35, 55, 80, 105 } : new[] { 20, 35, 55 };
            foreach (int row in rows) {
                foreach (int column in new[] { 110, 130, 155 }) {
                    OfficeColor pixel = raster.GetPixel(column, row);
                    Assert.NotEqual(OfficeColor.White, pixel);
                    Assert.Equal(control.GetPixel(column, pageIndex * 100 + row), pixel);
                }
            }
        }
        paged.RequireNoLoss();
        continuous.RequireNoLoss();
    }

    [Fact]
    public void HtmlGradientImageFragments_ParallelAutoHeightItemsRetainLegalContentCuts() {
        string first = string.Concat(Enumerable.Range(1, 7).Select(index => "<p>PAINT-" + index.ToString("D2") + "</p>"));
        string second = string.Concat(Enumerable.Range(1, 7).Select(index => "<p>SIDE-" + index.ToString("D2") + "</p>"));
        string html = GradientImageFragmentSource(
            "<p>PARALLEL-LEAD</p><section class='parallel'><div class='painted-parallel'>" + first
            + "</div><div class='plain-parallel'>" + second + "</div></section><p>PARALLEL-AFTER</p>",
            ".parallel{display:flex;width:180px;align-items:flex-start}.parallel p{height:20px;line-height:20px}"
            + ".painted-parallel{box-sizing:content-box;width:78px;padding:4px;border:2px solid green;"
            + GradientImageFragmentPaint("linear") + "}"
            + ".plain-parallel{box-sizing:border-box;width:90px;padding-top:6px}");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, GradientImageFragmentOptions(HtmlRenderMode.Paged));

        Assert.Equal(2, rendered.Pages.Count);
        AssertGradientImageFragmentMarkersOnce(rendered, "PARALLEL-LEAD", "PARALLEL-AFTER");
        for (int index = 1; index <= 7; index++)
            AssertGradientImageFragmentMarkersOnce(rendered, "PAINT-" + index.ToString("D2"), "SIDE-" + index.ToString("D2"));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment);
        Assert.All(rendered.Pages, page => {
            OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(page.CreateDrawing());
            Assert.NotEqual(OfficeColor.White, raster.GetPixel(80, 35));
        });
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlGradientImageFragments_ReplacedSvgImagesStillMoveWholeToTheNextPage() {
        string image = GradientImageFragmentSvgData(
            "<svg xmlns='http://www.w3.org/2000/svg' width='40' height='80' viewBox='0 0 40 80'>"
            + "<rect width='40' height='80' fill='blue'/></svg>");
        string html = GradientImageFragmentSource(
            "<div style='height:40px'>ATOMIC-LEAD</div><img src='" + image
            + "' style='display:block;margin:0;width:40px;height:80px'><p>ATOMIC-AFTER</p>", "");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, GradientImageFragmentOptions(HtmlRenderMode.Paged));

        Assert.Equal(2, rendered.Pages.Count);
        AssertGradientImageFragmentMarkersOnce(rendered, "ATOMIC-LEAD", "ATOMIC-AFTER");
        OfficeRasterImage first = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        OfficeRasterImage second = OfficeDrawingRasterRenderer.Render(rendered.Pages[1].CreateDrawing());
        Assert.Equal(OfficeColor.White, first.GetPixel(20, 75));
        Assert.Equal(OfficeColor.Blue, second.GetPixel(20, 15));
        Assert.Equal(OfficeColor.Blue, second.GetPixel(20, 85));
        Assert.Equal(OfficeColor.White, second.GetPixel(55, 85));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment);
        rendered.RequireNoLoss();
    }

    private static string GradientImageFragmentSource(string content, string rules) =>
        "<style>html,body,p{margin:0;padding:0;font:10px/10px '" + ManagedTextShapingTestAssets.FamilyName
        + "';font-kerning:none}body{width:180px}p{height:10px}" + rules + "</style>" + content;

    private static string GradientImageFragmentPaint(string family) {
        string image = family switch {
            "linear" => "linear-gradient(to bottom,red 0px 9px,blue 9px 18px)",
            "radial" => "radial-gradient(circle at center,red 0px 4px,blue 4px 8px)",
            "conic" => "conic-gradient(from 45deg,red 0deg 180deg,blue 180deg 360deg)",
            "svg" => "url(" + GradientImageFragmentSvgData(
                "<svg xmlns='http://www.w3.org/2000/svg' width='20' height='18' viewBox='0 0 20 18'>"
                + "<rect width='20' height='9' fill='red'/><rect y='9' width='20' height='9' fill='blue'/></svg>") + ")",
            _ => throw new ArgumentOutOfRangeException(nameof(family))
        };
        return "background-image:" + image + ";background-size:20px 18px;background-position:-3px -5px;"
            + "background-repeat:repeat;background-origin:content-box;background-clip:border-box;";
    }

    private static string GradientImageFragmentSvgData(string svg) =>
        "data:image/svg+xml;base64," + Convert.ToBase64String(Encoding.UTF8.GetBytes(svg));

    private static HtmlRenderOptions GradientImageFragmentOptions(HtmlRenderMode mode) {
        string family = ManagedTextShapingTestAssets.FamilyName;
        var options = new HtmlRenderOptions {
            Mode = mode,
            PageSize = new OfficePageSize(global::OfficeIMO.OfficeLength.Inches(200D / HtmlRenderOptions.CssPixelsPerInch), global::OfficeIMO.OfficeLength.Inches(120D / HtmlRenderOptions.CssPixelsPerInch)),
            Margins = HtmlRenderMargins.All(10D),
            HonorCssPageRules = false,
            ViewportWidth = 200D,
            ViewportHeight = 240D,
            AllowSystemFontFallback = false,
            DefaultFontFamily = family,
            UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser,
            BackgroundColor = OfficeColor.White
        };
        options.Fonts.Add(family, ManagedTextShapingTestAssets.CreateFontWithScalarAdvances(
            Enumerable.Range(32, 95).ToDictionary(scalar => scalar, _ => (ushort)500)));
        return options;
    }

    private static void AssertGradientImageFragmentMarkersOnce(HtmlRenderDocument rendered, params string[] markers) {
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderText>().ToArray();
        foreach (string marker in markers) Assert.Single(text, visual => visual.Text == marker);
    }
}

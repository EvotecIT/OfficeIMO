using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("root", HtmlRenderUserAgentStyleMode.Document)]
    [InlineData("root", HtmlRenderUserAgentStyleMode.Browser)]
    [InlineData("nested", HtmlRenderUserAgentStyleMode.Browser)]
    [InlineData("contents", HtmlRenderUserAgentStyleMode.Browser)]
    [InlineData("root-contents", HtmlRenderUserAgentStyleMode.Browser)]
    public void HtmlMargins_NegativeNetBlockAdvanceMovesFollowingFlowWithoutAnExtraPage(
        string route, HtmlRenderUserAgentStyleMode defaults) {
        string content = NegativeFlowBoxes("margin-top:-45px");
        if (route == "nested") content = "<section>" + content + "</section>";
        if (route == "contents") content = "<section style='display:contents'>" + content + "</section>";
        foreach (HtmlRenderMode mode in new[] { HtmlRenderMode.Continuous, HtmlRenderMode.Paged }) {
            HtmlRenderDocument rendered = RenderNegativeFlow(content, mode, defaults,
                route == "root-contents" ? "body{display:contents}" : "");
            Assert.Single(rendered.Pages);
            HtmlRenderShape overlap = NegativeFlowShape(rendered, "div#overlap");
            HtmlRenderShape after = NegativeFlowShape(rendered, "div#after");
            Assert.Equal(35D, overlap.Y, 6);
            Assert.Equal(25D, overlap.Height, 6);
            Assert.Equal(60D, after.Y, 6);
            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
            Assert.Equal(OfficeColor.Lime, image.GetPixel(10, 40));
            Assert.Equal(OfficeColor.Blue, image.GetPixel(10, 70));
        }
    }

    [Fact]
    public void HtmlMargins_ZeroNetBlockAdvanceDoesNotAddAnEpsilonToTheNextSibling() {
        HtmlRenderDocument rendered = RenderNegativeFlow(NegativeFlowBoxes("margin-top:-25px"),
            HtmlRenderMode.Paged, HtmlRenderUserAgentStyleMode.Document);
        Assert.Single(rendered.Pages);
        Assert.Equal(55D, NegativeFlowShape(rendered, "div#overlap").Y, 6);
        Assert.Equal(80D, NegativeFlowShape(rendered, "div#after").Y, 6);
    }

    [Fact]
    public void HtmlMargins_NegativeFirstChildMarginRetainsItsSignedAdvanceWhenCollapsedThroughAParent() {
        const string content = "<div style='height:80px'></div><section>"
            + "<div id='overlap' style='height:25px;margin-top:-45px;background:#00ff00'></div>"
            + "</section><div id='after' style='height:20px;background:blue'></div>";
        HtmlRenderDocument rendered = RenderNegativeFlow(content, HtmlRenderMode.Paged,
            HtmlRenderUserAgentStyleMode.Browser);
        Assert.Single(rendered.Pages);
        Assert.Equal(35D, NegativeFlowShape(rendered, "div#overlap").Y, 6);
        Assert.Equal(60D, NegativeFlowShape(rendered, "div#after").Y, 6);
    }

    [Fact]
    public void HtmlMargins_NegativeBottomMarginMovesFlowWithoutMovingOrShorteningThePaintedBox() {
        HtmlRenderDocument rendered = RenderNegativeFlow(NegativeFlowBoxes("margin-bottom:-45px"),
            HtmlRenderMode.Continuous, HtmlRenderUserAgentStyleMode.Browser);
        HtmlRenderShape overlap = NegativeFlowShape(rendered, "div#overlap");
        Assert.Equal(80D, overlap.Y, 6);
        Assert.Equal(25D, overlap.Height, 6);
        Assert.Equal(60D, NegativeFlowShape(rendered, "div#after").Y, 6);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        Assert.Equal(OfficeColor.Blue, image.GetPixel(10, 70));
        Assert.Equal(OfficeColor.Lime, image.GetPixel(10, 90));
    }

    [Fact]
    public void HtmlMargins_BreakAfterANegativeNetBlockUsesItsFollowingSiblingBoundary() {
        HtmlRenderDocument rendered = RenderNegativeFlow(
            NegativeFlowBoxes("margin-top:-45px;break-after:page"), HtmlRenderMode.Paged,
            HtmlRenderUserAgentStyleMode.Browser);
        Assert.Equal(2, rendered.Pages.Count);
        Assert.Equal(35D, NegativeFlowShape(rendered, "div#overlap").Y, 6);
        HtmlRenderPage followingPage = Assert.Single(rendered.Pages, page =>
            EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderShape>()
                .Any(shape => shape.Source == "div#after" && shape.Shape.FillColor.HasValue));
        Assert.Equal(2, followingPage.PageNumber);
        Assert.Equal(0D, NegativeFlowShape(rendered, "div#after").Y, 6);
    }

    [Theory]
    [InlineData(false, 1120D)]
    [InlineData(true, 1085D)]
    public void HtmlMargins_FullPageHeightRetainsPaintBeyondTheSignedCursorAndHonoursOverflowClipping(
        bool clipped, double expectedHeight) {
        string html = "<!doctype html><style>html,body{margin:0;padding:0;font-size:0;line-height:0}</style>"
            + "<div style='height:1060px'></div>"
            + "<div style='height:25px;margin-bottom:-45px;overflow:" + (clipped ? "hidden" : "visible") + "'>"
            + "<div id='paint-tail' style='height:60px;background:#00ff00'></div></div>"
            + "<div id='after' style='height:20px;background:blue'></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous, ViewportWidth = 120D, ViewportHeight = 1000D,
            Margins = HtmlRenderMargins.All(0D), UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser,
            AllowSystemFontFallback = false
        });
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        Assert.Equal(1040D, NegativeFlowShape(rendered, "div#after").Y, 6);
        Assert.Equal(expectedHeight, page.Height, 6);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(page.CreateDrawing());
        Assert.Equal((int)expectedHeight, image.Height);
        Assert.Equal(OfficeColor.Blue, image.GetPixel(10, 1050));
        Assert.Equal(OfficeColor.Lime, image.GetPixel(10, (int)expectedHeight - 1));
    }

    [Fact]
    public void HtmlMargins_NegativeNetAdvanceInsideAnOrdinaryColumnRetainsOriginalFlowAndPaint() {
        string content = "<section style='width:220px;height:100px;column-count:2;column-gap:20px;column-fill:auto'>"
            + NegativeFlowBoxes("margin-top:-45px") + "</section>";
        HtmlRenderDocument rendered = RenderNegativeFlow(content, HtmlRenderMode.Paged,
            HtmlRenderUserAgentStyleMode.Browser);
        Assert.Single(rendered.Pages);
        HtmlRenderShape overlap = NegativeFlowShape(rendered, "div#overlap");
        HtmlRenderShape after = NegativeFlowShape(rendered, "div#after");
        Assert.Equal(0D, overlap.X, 6);
        Assert.Equal(35D, overlap.Y, 6);
        Assert.Equal(0D, after.X, 6);
        Assert.Equal(60D, after.Y, 6);
    }

    [Fact]
    public void HtmlMargins_NegativeSiblingAfterADescendantSpannerKeepsTheOriginalSourceOnce() {
        const string content = "<section style='width:300px;column-count:2;column-gap:20px;column-fill:auto'>"
            + "<div style='opacity:.5;background:red'><p>P1</p><p>P2</p>"
            + "<div style='column-span:all'><p><a href='#after'>SPAN</a></p></div>"
            + "<p>T1</p><p>T2</p></div></section>"
            + "<div id='overlap' style='height:25px;margin-top:-45px;background:#00ff00'>OVERLAP</div>"
            + "<p id='after'>AFTER</p>";
        HtmlRenderDocument rendered = RenderNegativeFlow(content, HtmlRenderMode.Paged,
            HtmlRenderUserAgentStyleMode.Browser,
            "html,body{font:14px/20px Arial}p{margin:0;font:inherit;orphans:1;widows:1}");
        Assert.Single(rendered.Pages);
        Assert.Equal(35D, NegativeFlowShape(rendered, "div#overlap").Y, 4);
        Assert.Equal(60D, DescendantSpannerText(rendered, "AFTER").Y, 4);
        Assert.Equal("#after", DescendantSpannerText(rendered, "SPAN").LinkUri);
        AssertDescendantSpannerSourceOnce(rendered, "P1", "P2", "SPAN", "T1", "T2", "OVERLAP", "AFTER");
    }

    private static string NegativeFlowBoxes(string margin) =>
        "<div style='height:80px'></div>"
        + "<div id='overlap' style='height:25px;background:#00ff00;" + margin + "'></div>"
        + "<div id='after' style='height:20px;background:blue'></div>";

    private static HtmlRenderShape NegativeFlowShape(HtmlRenderDocument rendered, string source) =>
        Assert.Single(rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderShape>(), shape => shape.Source == source && shape.Shape.FillColor.HasValue);

    private static HtmlRenderDocument RenderNegativeFlow(string content, HtmlRenderMode mode,
        HtmlRenderUserAgentStyleMode defaults, string css = "") =>
        HtmlRenderTestDriver.Render("<!doctype html><style>html,body{margin:0;padding:0;font-size:0;line-height:0}"
            + css + "</style>" + content, new HtmlRenderOptions {
                Mode = mode, ViewportWidth = 700D, ViewportHeight = 140D,
                PageSize = new OfficePageSize(global::OfficeIMO.OfficeLength.Inches(700D / 96D), global::OfficeIMO.OfficeLength.Inches(100D / 96D)),
                Margins = HtmlRenderMargins.All(0D), HonorCssPageRules = false,
                UserAgentStyles = defaults, AllowSystemFontFallback = false
            });
}

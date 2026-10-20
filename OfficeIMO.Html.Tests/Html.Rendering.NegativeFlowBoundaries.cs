using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlMargins_ZeroNetOrdinaryBoxAtThePageEndDoesNotCreateAPaintCarrierPage(bool contentsRoot) {
        const string content = "<div style='height:100px'></div>"
            + "<div id='overlap' style='height:25px;margin-top:-25px;background:#00ff00'></div>";
        HtmlRenderDocument rendered = RenderNegativeFlow(content, HtmlRenderMode.Paged,
            HtmlRenderUserAgentStyleMode.Browser, contentsRoot ? "body{display:contents}" : "");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        Assert.Equal(100D, page.Height, 4);
        Assert.Equal(75D, NegativeFlowShape(rendered, "div#overlap").Y, 4);
        Assert.Equal(OfficeColor.Lime, OfficeDrawingRasterRenderer.Render(
            page.CreateDrawing()).GetPixel(10, 99));
    }

    [Fact]
    public void HtmlMargins_SignedColumnSiblingsRetainTheFiniteChildsIndependentOverflowProjection() {
        const string content = "<section style='width:220px;height:60px;column-count:2;column-gap:20px;column-fill:auto'>"
            + "<div style='height:20px'><p>L1</p><p>L2</p><p>L3</p><p>L4</p><p>L5</p>"
            + "<p><a href='#tail'>L6</a></p></div>"
            + "<div style='height:10px;margin-top:-20px'></div><p id='tail'>TAIL</p></section>";
        HtmlRenderDocument rendered = RenderNegativeFlow(content, HtmlRenderMode.Paged,
            HtmlRenderUserAgentStyleMode.Browser,
            "html,body{font:14px/20px Arial}p{margin:0;font:inherit;orphans:1;widows:1}");
        Assert.Single(rendered.Pages);
        Assert.Equal(120D, DescendantSpannerText(rendered, "L4").X, 4);
        Assert.Equal(0D, DescendantSpannerText(rendered, "L4").Y, 4);
        Assert.Equal(10D, DescendantSpannerText(rendered, "TAIL").Y, 4);
        Assert.Equal("#tail", DescendantSpannerText(rendered, "L6").LinkUri);
        AssertDescendantSpannerSourceOnce(rendered, "L1", "L2", "L3", "L4", "L5", "L6", "TAIL");
    }

    [Fact]
    public void HtmlMargins_RewindingColumnFlowStillAvoidsAnEarlierFloatBand() {
        const string content = "<section style='width:220px;height:120px;column-count:2;column-gap:20px;column-fill:auto'>"
            + "<div style='float:left;width:40px;height:60px;background:red'></div>"
            + "<div style='height:80px'></div><div style='height:10px;margin-top:-60px'></div>"
            + "<div id='bfc' style='display:flow-root;width:50px;height:10px;background:blue'></div></section>";
        HtmlRenderDocument rendered = RenderNegativeFlow(content, HtmlRenderMode.Continuous,
            HtmlRenderUserAgentStyleMode.Browser);
        HtmlRenderShape bfc = NegativeFlowShape(rendered, "div#bfc");
        Assert.Equal(40D, bfc.X, 4);
        Assert.Equal(30D, bfc.Y, 4);
        Assert.Equal(OfficeColor.Blue, OfficeDrawingRasterRenderer.Render(
            rendered.Pages[0].CreateDrawing()).GetPixel(45, 35));
    }

    [Fact]
    public void HtmlMargins_ContentsRootRetainsOriginalSourceWhileReflowingOntoANarrowerPage() {
        string[] words = Enumerable.Range(1, 40).Select(index => "WORD" + index.ToString("D2")).ToArray();
        string html = "<!doctype html><style>html,body{margin:0;padding:0;font:14px/20px Arial}"
            + "body{display:contents}p{margin:0;orphans:1;widows:1}"
            + "@page{size:700px 140px;margin:20px}@page:left{size:260px 140px;margin:20px}</style>"
            + "<div style='height:80px'>FIRST</div><div style='height:25px;margin-top:-45px'>OVER</div>"
            + "<p>" + string.Join(" ", words) + "</p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser,
            AllowSystemFontFallback = false
        });
        Assert.True(rendered.Pages.Count > 1);
        Assert.Equal(260D, rendered.Pages[1].Width, 4);
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.PagePseudoGeometryPending);
        string[] actual = rendered.Text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(words.OrderBy(word => word, StringComparer.Ordinal), actual
            .Where(word => word.StartsWith("WORD", StringComparison.Ordinal)).OrderBy(word => word, StringComparer.Ordinal));
        HtmlRenderText[] narrowPageText = EnumerateRenderVisuals(rendered.Pages[1].Scene)
            .OfType<HtmlRenderText>().ToArray();
        Assert.NotEmpty(narrowPageText);
        Assert.All(narrowPageText, text =>
            Assert.True(text.X + text.Width <= 240D + 0.001D));
    }

    [Fact]
    public void HtmlMargins_TableCellRetainsItsPositivePaintTailBeyondSignedNormalFlow() {
        const string content = "<table style='width:80px'><tr><td>"
            + "<div style='height:80px'>FIRST</div>"
            + "<div id='overlap' style='height:25px;margin-bottom:-45px;background:#00ff00'>OVER</div>"
            + "<div id='after' style='height:20px;background:blue'>TAIL</div></td></tr></table>";
        HtmlRenderDocument rendered = RenderNegativeFlow(content, HtmlRenderMode.Paged,
            HtmlRenderUserAgentStyleMode.Browser,
            "html,body{font:14px/20px Arial}table,td{margin:0;border:0;padding:0;border-spacing:0}td{vertical-align:top}");
        Assert.Equal(2, rendered.Pages.Count);
        Assert.Equal(60D, NegativeFlowShape(rendered, "div#after").Y, 4);
        var greenRows = new List<int>();
        foreach (HtmlRenderPage page in rendered.Pages) {
            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(page.CreateDrawing());
            int count = 0;
            for (int y = 0; y < image.Height; y++)
                if (image.GetPixel(70, y) == OfficeColor.Lime) count++;
            greenRows.Add(count);
        }
        Assert.Equal(new[] { 20, 5 }, greenRows);
        Assert.Equal(25, greenRows.Sum());
        AssertDescendantSpannerSourceOnce(rendered, "FIRST", "OVER", "TAIL");
    }
}

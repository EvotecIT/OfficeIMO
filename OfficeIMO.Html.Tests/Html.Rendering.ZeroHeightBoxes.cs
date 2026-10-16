using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlZeroHeightBoxTests {
    private static readonly OfficeColor TailColor = OfficeColor.FromRgb(208, 108, 36);
    private const string Tail = "<div id='tail' style='width:120px;height:2px;background:#d06c24'></div>";

    [Theory]
    [InlineData("div")]
    [InlineData("table")]
    [InlineData("caption")]
    [InlineData("td")]
    public void ZeroLayoutHeightRetainsExpandedPaintWithoutAdvancingFollowingFlow(string owner) {
        string box = owner switch {
            "div" => "<div id='zero' class='zero'></div>",
            "table" => "<table id='zero' class='zero'></table>",
            "caption" => "<table><caption id='zero' class='zero'></caption>"
                + "<tr><td style='width:100px'></td></tr></table>",
            _ => "<table><tr><td id='zero' class='zero'></td></tr></table>"
        };
        HtmlRenderPage page = Render(box + Tail, ".zero{width:100px;height:0;"
            + "box-shadow:0 0 1px 5px red;outline:2px solid blue}");
        HtmlRenderShape tail = FilledShape(page, "div#tail");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(page.CreateDrawing());

        // Native controls isolate zero flow and expanded paint. Samples avoid
        // antialias edges and include outline overlap with the sibling background.
        Assert.Equal(16D, tail.Y, 6);
        Assert.Equal(2D, tail.Height, 6);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(32, 13));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(32, 14));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(32, 16));
        Assert.Equal(TailColor, raster.GetPixel(128, 16));
    }

    [Theory]
    [InlineData(1, false, 100D, 0D)]
    [InlineData(2, false, 50D, 50D)]
    [InlineData(2, true, 0D, 100D)]
    public void EmptyPercentageRowsRetainTheAuthoredTableMinimum(
        int rowCount, bool secondAuto, double firstHeight, double secondHeight) {
        string rows = "<tr style='height:0%'><td id='first' style='background:lime'></td></tr>";
        if (rowCount == 2) rows += "<tr" + (secondAuto ? "" : " style='height:0%'")
            + "><td id='second' style='background:blue'></td></tr>";
        HtmlRenderPage page = Render("<table id='grid' style='width:100px;height:100px;"
            + "border-collapse:collapse;background:#eee'>" + rows + "</table>" + Tail);
        HtmlRenderShape grid = FilledShape(page, "table#grid");
        HtmlRenderShape tail = FilledShape(page, "div#tail");

        Assert.Equal(16D, grid.Y, 6);
        Assert.Equal(100D, grid.Height, 6);
        Assert.Equal(116D, tail.Y, 6);
        if (firstHeight > 0D) {
            HtmlRenderShape first = FilledShape(page, "td#first");
            Assert.Equal(16D, first.Y, 6);
            Assert.Equal(firstHeight, first.Height, 6);
        } else {
            // A zero share beside an auto row has no background paint area.
            Assert.DoesNotContain(page.Visuals.OfType<HtmlRenderShape>(), shape =>
                shape.Source == "td#first" && shape.Shape.FillColor.HasValue);
        }
        if (rowCount == 2) {
            HtmlRenderShape second = FilledShape(page, "td#second");
            Assert.Equal(16D + firstHeight, second.Y, 6);
            Assert.Equal(secondHeight, second.Height, 6);
        }
    }

    private static HtmlRenderShape FilledShape(HtmlRenderPage page, string source) =>
        Assert.Single(page.Visuals.OfType<HtmlRenderShape>(), shape =>
            shape.Source == source && shape.Shape.FillColor.HasValue);

    private static HtmlRenderPage Render(string content, string extraCss = "") =>
        Assert.Single(HtmlRenderTestDriver.Render("<!doctype html><style>"
            + "html,body{margin:0;padding:0;font-size:0;line-height:0;background:white}"
            + "body{padding:16px}table,caption,td,tr{border:0;padding:0;margin:0;border-spacing:0}"
            + extraCss + "</style>" + content, new HtmlRenderOptions {
                ViewportWidth = 240D, ViewportHeight = 180D,
                Margins = HtmlRenderMargins.All(0D),
                UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser,
                AllowSystemFontFallback = false
            }).Pages);
}

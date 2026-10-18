using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("hidden", false)]
    [InlineData("clip", false)]
    [InlineData("hidden", true)]
    [InlineData("clip", true)]
    public void HtmlMargins_CellOverflowUsesItsActualPaddingEdgeWithoutAddingAPaintTailPage(
        string overflow, bool decorated) {
        string decoration = decorated ? "padding:5px;border:2px solid red;outline:2px solid purple;" : "";
        string content = "<table style='width:120px'><tr><td id='cell' style='height:40px;background:#eeeeee;"
            + decoration + "overflow:" + overflow + "'>" + ClippedTableCellSource + "</td></tr></table>";
        foreach (HtmlRenderMode mode in new[] { HtmlRenderMode.Continuous, HtmlRenderMode.Paged }) {
            HtmlRenderDocument rendered = RenderNegativeFlow(content, mode, HtmlRenderUserAgentStyleMode.Browser,
                ClippedTableCellCss);
            HtmlRenderPage page = Assert.Single(rendered.Pages);
            HtmlRenderClipGroup clip = Assert.Single(EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderClipGroup>(),
                group => group.Source == "td#cell");
            Assert.Equal(decorated ? 2D : 0D, clip.ClipX, 4);
            Assert.Equal(decorated ? 2D : 0D, clip.ClipY, 4);
            Assert.Equal(decorated ? 90D : 80D, clip.ClipHeight, 4);
            Assert.Equal(decorated ? 67D : 60D, NegativeFlowShape(rendered, "div#after").Y, 4);
            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(page.CreateDrawing());
            Assert.Equal(OfficeColor.Blue, image.GetPixel(70, decorated ? 77 : 70));
            int greenRows = Enumerable.Range(0, image.Height)
                .Count(y => image.GetPixel(70, y) == OfficeColor.Lime);
            Assert.Equal(decorated ? 5 : 0, greenRows);
            if (decorated) {
                Assert.Equal(OfficeColor.Red, image.GetPixel(70, 0));
                Assert.Equal(OfficeColor.FromRgb(238, 238, 238), image.GetPixel(70, 3));
                Assert.Equal(OfficeColor.Red, image.GetPixel(70, 93));
                Assert.Equal(OfficeColor.Purple, image.GetPixel(70, 95));
            }
            Assert.Equal("#after", DescendantSpannerText(rendered, "OVER").LinkUri);
            AssertDescendantSpannerSourceOnce(rendered, "FIRST", "OVER", "TAIL");
        }
    }

    [Theory]
    [InlineData("hidden")]
    [InlineData("clip")]
    public void HtmlMargins_CellOverflowUsesTheWholeRowspanInsteadOfItsFirstRowHeight(string overflow) {
        string content = "<table style='width:160px'><tr><td id='cell' rowspan='2' style='height:40px;overflow:"
            + overflow + "'>" + ClippedTableCellSource
            + "</td><td style='width:40px;height:40px'></td></tr>"
            + "<tr><td style='width:40px;height:40px'></td></tr></table>";
        foreach (HtmlRenderMode mode in new[] { HtmlRenderMode.Continuous, HtmlRenderMode.Paged }) {
            HtmlRenderDocument rendered = RenderNegativeFlow(content, mode, HtmlRenderUserAgentStyleMode.Browser,
                ClippedTableCellCss);
            HtmlRenderPage page = Assert.Single(rendered.Pages);
            HtmlRenderClipGroup clip = Assert.Single(EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderClipGroup>(),
                group => group.Source == "td#cell");
            Assert.Equal(80D, clip.ClipHeight, 4);
            Assert.Equal(60D, NegativeFlowShape(rendered, "div#after").Y, 4);
            Assert.Equal(OfficeColor.Blue, OfficeDrawingRasterRenderer.Render(page.CreateDrawing()).GetPixel(70, 70));
            AssertDescendantSpannerSourceOnce(rendered, "FIRST", "OVER", "TAIL");
        }
    }

    private const string ClippedTableCellCss = "html,body{font:14px/20px Arial}"
        + "table,td{margin:0;border:0;padding:0;border-spacing:0}td{vertical-align:top}a{color:inherit;text-decoration:none}";

    private const string ClippedTableCellSource = "<div style='height:80px'>FIRST</div>"
        + "<div style='height:25px;margin-bottom:-45px;background:#00ff00'><a href='#after'>OVER</a></div>"
        + "<div id='after' style='height:20px;background:blue'>TAIL</div>";
}

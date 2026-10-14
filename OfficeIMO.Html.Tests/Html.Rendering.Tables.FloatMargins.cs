using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("left", "0 auto", 350D)]
    [InlineData("right", "0 auto", 50D)]
    [InlineData("left", "0 20px 0 auto", 380D)]
    [InlineData("right", "0 20px 0 auto", 80D)]
    [InlineData("left", "0 auto 0 20px", 320D)]
    [InlineData("right", "0 auto 0 20px", 20D)]
    public void HtmlTables_AutoMarginsResolveWithinTheFreeFloatBand(string side, string margins, double expectedX) {
        foreach (bool pagedReflow in new[] { false, true }) {
            HtmlRenderDocument rendered = RenderTableFloatMargins(side,
                "<table id='grid' style='width:200px;margin:" + margins
                + ";background:blue'><tr><td></td></tr></table>", pagedReflow);

            AssertTableFloatBand(rendered, side, "table#grid", expectedX, 200D);
        }
    }

    [Theory]
    [InlineData("table", "width:100px;table-layout:fixed", "<colgroup><col span='2' style='width:100px'></colgroup><tr><td></td><td></td></tr>", 200D)]
    [InlineData("table", "width:200px;height:40px", "", 200D)]
    [InlineData("table", "", "<tr><td style='width:200px'></td></tr>", 200D)]
    [InlineData("div", "display:table;width:200px", "<div style='display:table-row'><div style='display:table-cell;height:40px'></div></div>", 200D)]
    public void HtmlTables_UsedWidthAndEmptyOrCssTableMarginsPreserveFloatAvoidance(
        string tag, string sizing, string content, double expectedWidth) {
        foreach (bool pagedReflow in new[] { false, true }) {
            HtmlRenderDocument rendered = RenderTableFloatMargins("left", "<" + tag
                + " id='grid' style='" + sizing + ";margin:0 auto;background:blue'>"
                + content + "</" + tag + ">", pagedReflow);

            AssertTableFloatBand(rendered, "left", tag + "#grid", 350D, expectedWidth);
        }
    }

    [Theory]
    [InlineData("left", 350D)]
    [InlineData("right", 50D)]
    public void HtmlTables_PagedHeightRecheckAvoidsAFloatStartingInsideTheTable(string side, double expectedX) {
        HtmlRenderDocument rendered = RenderTableFloatMargins(side,
            "<table id='grid' style='width:200px;margin:0 auto;background:blue'><tr><td></td></tr></table>",
            pagedReflow: true, delayedFloat: true);

        AssertTableFloatBand(rendered, side, "table#grid", expectedX, 200D, sameTop: false);
        HtmlRenderShape table = FixedMinimumShape(rendered, "table#grid");
        HtmlRenderShape floating = FixedMinimumShape(rendered, "span#float");
        Assert.Equal(0D, table.Y, 3);
        Assert.Equal(20D, floating.Y, 3);
        Assert.True(floating.Y < table.Y + table.Height);
    }

    [Theory]
    [InlineData("left")]
    [InlineData("right")]
    public void HtmlTables_FixedExpandedWidthMovesBelowAFloatWhenTheFreeBandIsTooNarrow(string side) {
        foreach (bool pagedReflow in new[] { false, true }) {
            HtmlRenderDocument rendered = RenderTableFloatMargins(side,
                "<table id='grid' style='width:100px;table-layout:fixed;margin:0 auto;background:blue'>"
                + "<colgroup><col span='2' style='width:200px'></colgroup><tr><td></td><td></td></tr></table>", pagedReflow);
            HtmlRenderShape table = FixedMinimumShape(rendered, "table#grid");
            HtmlRenderShape floating = FixedMinimumShape(rendered, "span#float");

            Assert.Equal(400D, table.Width, 3);
            Assert.Equal(100D, table.X, 3);
            Assert.Equal(floating.Y + floating.Height, table.Y, 3);
        }
    }

    private static HtmlRenderDocument RenderTableFloatMargins(string side, string table, bool pagedReflow, bool delayedFloat = false) {
        bool parallelColumn = pagedReflow && !delayedFloat;
        string html = "<!doctype html><style>html,body{margin:0;padding:0;font-size:0;line-height:0}"
            + "table,td{padding:0;border:0;border-spacing:0}td{height:40px}</style>"
            // A parallel flex column reports its floats through the active-float
            // path, while continuous flow uses the shared inline float context.
            + (parallelColumn ? "<main style='display:flex'><article style='width:600px'><div style='height:260px'></div><p style='margin:0'>" : string.Empty)
            + (delayedFloat ? "<div style='height:280px'></div>" : string.Empty)
            + "<span id='float' style='float:" + side + ";width:300px;height:120px;background:red'></span>"
            + (parallelColumn ? "</p>" : string.Empty) + table
            + (parallelColumn ? "</article></main>" : string.Empty);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = pagedReflow ? HtmlRenderMode.Paged : HtmlRenderMode.Continuous,
            UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser,
            ViewportWidth = 600D, ViewportHeight = 300D,
            PageSize = new OfficePageSize(600D / HtmlRenderOptions.CssPixelsPerInch, 300D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false, Margins = HtmlRenderMargins.All(0D), AllowSystemFontFallback = false,
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        });
        rendered.RequireNoLoss();
        Assert.Empty(rendered.Text);
        Assert.Empty(rendered.Fonts.Faces);
        return rendered;
    }

    private static void AssertTableFloatBand(HtmlRenderDocument rendered, string side, string source, double x, double width, bool sameTop = true) {
        HtmlRenderPage page = Assert.Single(rendered.Pages, candidate =>
            FixedMinimumShapes(candidate.Scene).Any(shape => shape.Source == source));
        HtmlRenderShape table = Assert.Single(FixedMinimumShapes(page.Scene), shape => shape.Source == source);
        HtmlRenderShape floating = Assert.Single(FixedMinimumShapes(page.Scene), shape => shape.Source == "span#float");
        Assert.Equal(x, table.X, 3);
        Assert.Equal(width, table.Width, 3);
        if (sameTop) Assert.Equal(floating.Y, table.Y, 3);
        if (side == "left") Assert.True(table.X >= floating.X + floating.Width - 0.001D);
        else Assert.True(table.X + table.Width <= floating.X + 0.001D);
    }
}

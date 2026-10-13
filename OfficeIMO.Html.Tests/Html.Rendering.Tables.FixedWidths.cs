using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("column")]
    [InlineData("cell")]
    public void HtmlTables_FixedAbsoluteTracksExpandPastTheAuthoredTableWidth(string declaration) {
        string columns = declaration == "column" ? "<colgroup><col style='width:80px'><col style='width:80px'></colgroup>" : "";
        string cellWidth = declaration == "cell" ? "width:80px;" : "";
        HtmlRenderDocument rendered = RenderFixedMinimum("<table id='grid' style='width:100px;table-layout:fixed;background:lime'>"
            + columns + "<tr><td id='left' style='" + cellWidth + "background:red'></td><td id='right' style='"
            + cellWidth + "background:blue'></td></tr></table><div id='following' style='width:5px;height:4px;background:black'></div>");

        Assert.Equal(160D, FixedMinimumShape(rendered, "table#grid").Width, 3);
        Assert.Equal(80D, FixedMinimumShape(rendered, "td#left").Width, 3);
        Assert.Equal(80D, FixedMinimumShape(rendered, "td#right").Width, 3);
        Assert.Equal(80D, FixedMinimumShape(rendered, "td#right").X, 3);
        Assert.Equal(12D, FixedMinimumShape(rendered, "div#following").Y, 3);
    }

    [Fact]
    public void HtmlTables_FixedColumnDeclarationsTakePrecedenceOverFirstRowWidths() {
        HtmlRenderDocument rendered = RenderFixedMinimum("<table id='grid' style='width:100px;table-layout:fixed;background:lime'>"
            + "<colgroup><col span='2' style='width:80px'></colgroup><tr>"
            + "<td id='left' style='width:200px;background:red'></td><td id='right' style='width:200px;background:blue'></td></tr></table>");

        Assert.Equal(160D, FixedMinimumShape(rendered, "table#grid").Width, 3);
        Assert.Equal(80D, FixedMinimumShape(rendered, "td#left").Width, 3);
        Assert.Equal(80D, FixedMinimumShape(rendered, "td#right").Width, 3);
    }

    [Theory]
    [InlineData("column")]
    [InlineData("cell")]
    public void HtmlTables_FixedPercentageCompressionPreservesAbsoluteTracks(string declaration) {
        string columns = declaration == "column" ? "<colgroup><col style='width:80px'><col style='width:80%'></colgroup>" : "";
        string leftWidth = declaration == "cell" ? "width:80px;" : "";
        string rightWidth = declaration == "cell" ? "width:80%;" : "";
        HtmlRenderDocument rendered = RenderFixedMinimum("<table id='grid' style='width:100px;table-layout:fixed;background:lime'>"
            + columns + "<tr><td id='left' style='" + leftWidth + "background:red'></td><td id='right' style='"
            + rightWidth + "background:blue'></td></tr></table>");

        Assert.Equal(100D, FixedMinimumShape(rendered, "table#grid").Width, 3);
        Assert.Equal(80D, FixedMinimumShape(rendered, "td#left").Width, 3);
        Assert.Equal(20D, FixedMinimumShape(rendered, "td#right").Width, 3);
        Assert.Equal(80D, FixedMinimumShape(rendered, "td#right").X, 3);
    }

    [Fact]
    public void HtmlTables_FixedOvercommittedPercentagesStayWithinTheAuthoredTableWidth() {
        HtmlRenderDocument rendered = RenderFixedMinimum("<table id='grid' style='width:100px;table-layout:fixed;background:lime'>"
            + "<colgroup><col span='2' style='width:80%'></colgroup><tr>"
            + "<td id='left' style='background:red'></td><td id='right' style='background:blue'></td></tr></table>");

        Assert.Equal(100D, FixedMinimumShape(rendered, "table#grid").Width, 3);
        Assert.Equal(50D, FixedMinimumShape(rendered, "td#left").Width, 3);
        Assert.Equal(50D, FixedMinimumShape(rendered, "td#right").Width, 3);
    }

    [Fact]
    public void HtmlTables_FixedColspanHintsShareWidthOnlyWithUndeclaredColumns() {
        HtmlRenderDocument rendered = RenderFixedMinimum("<table id='grid' style='width:100px;table-layout:fixed;background:lime'>"
            + "<colgroup><col style='width:80px'><col></colgroup><tr><td id='spanning' colspan='2' style='width:200px;background:gray'></td></tr>"
            + "<tr><td id='left' style='background:red'></td><td id='right' style='background:blue'></td></tr></table>"
            + "<div id='following' style='width:5px;height:4px;background:black'></div>");

        Assert.Equal(180D, FixedMinimumShape(rendered, "table#grid").Width, 3);
        Assert.Equal(180D, FixedMinimumShape(rendered, "td#spanning").Width, 3);
        Assert.Equal(80D, FixedMinimumShape(rendered, "td#left").Width, 3);
        Assert.Equal(100D, FixedMinimumShape(rendered, "td#right").Width, 3);
        Assert.Equal(80D, FixedMinimumShape(rendered, "td#right").X, 3);
        Assert.Equal(24D, FixedMinimumShape(rendered, "div#following").Y, 3);
    }

    [Theory]
    [InlineData("0 auto", 70D)]
    [InlineData("0 20px 0 auto", 120D)]
    [InlineData("0 auto 0 20px", 20D)]
    public void HtmlTables_FixedExpansionResolvesAutoMarginsFromTheUsedBorderBox(string margins, double x) {
        HtmlRenderDocument rendered = RenderFixedMinimum("<table id='grid' style='width:100px;table-layout:fixed;background:lime;margin:"
            + margins + "'><colgroup><col span='2' style='width:80px'></colgroup><tr><td></td><td></td></tr></table>");

        HtmlRenderShape table = FixedMinimumShape(rendered, "table#grid");
        Assert.Equal(160D, table.Width, 3);
        Assert.Equal(x, table.X, 3);
    }

    [Fact]
    public void HtmlTables_FixedExpansionCountsSpacingAndTableInsetsOnce() {
        HtmlRenderDocument rendered = RenderFixedMinimum("<table id='grid' style='width:100px;table-layout:fixed;background:lime;"
            + "border-spacing:4px 0;padding:5px;border:2px solid black'><colgroup><col span='2' style='width:80px'></colgroup>"
            + "<tr><td id='left' style='background:red'></td><td id='right' style='background:blue'></td></tr></table>");

        Assert.Equal(186D, FixedMinimumShape(rendered, "table#grid", OfficeColor.FromRgb(0, 255, 0)).Width, 3);
        Assert.Equal(80D, FixedMinimumShape(rendered, "td#left").Width, 3);
        Assert.Equal(80D, FixedMinimumShape(rendered, "td#right").Width, 3);
        Assert.Equal(11D, FixedMinimumShape(rendered, "td#left").X, 3);
        Assert.Equal(95D, FixedMinimumShape(rendered, "td#right").X, 3);
    }

    [Theory]
    [InlineData("hidden")]
    [InlineData("auto")]
    public void HtmlTables_FixedOverflowClipUsesTheExpandedTableBorderBox(string overflow) {
        HtmlRenderDocument rendered = RenderFixedMinimum("<table id='grid' style='width:100px;table-layout:fixed;background:lime;overflow:"
            + overflow + "'><colgroup><col span='2' style='width:80px'></colgroup><tr>"
            + "<td id='left' style='background:red'></td><td id='right' style='background:blue'></td></tr></table>");

        HtmlRenderClipGroup clip = Assert.Single(EnumerateTablePaginationScene(rendered.Pages[0].Scene)
            .OfType<HtmlRenderClipGroup>(), group => group.Source == "table#grid");
        Assert.Equal(160D, clip.ClipWidth, 3);
        Assert.Equal(80D, FixedMinimumShape(rendered, "td#right").Width, 3);
        Assert.True(clip.ClipX + clip.ClipWidth >= FixedMinimumShape(rendered, "td#right").X + 80D);
    }

    [Fact]
    public void HtmlTables_FixedInlineTableAdvancesFollowingAtomicContentByItsUsedWidth() {
        HtmlRenderDocument rendered = RenderFixedMinimum("<table id='grid' style='display:inline-table;vertical-align:top;width:100px;"
            + "table-layout:fixed;background:lime'><colgroup><col span='2' style='width:80px'></colgroup><tr><td></td><td></td></tr></table>"
            + "<span id='after' style='display:inline-block;vertical-align:top;width:10px;height:12px;background:red'></span>");

        HtmlRenderShape table = FixedMinimumShape(rendered, "table#grid");
        HtmlRenderShape after = FixedMinimumShape(rendered, "span#after");
        Assert.Equal(160D, table.Width, 3);
        Assert.Equal(table.X + table.Width, after.X, 3);
        Assert.Equal(table.Y, after.Y, 3);
    }

    [Fact]
    public void HtmlTables_FixedNestedInlineTableContributesItsExpandedWidthToTheOuterCell() {
        HtmlRenderDocument rendered = RenderFixedMinimum("<table><tr><td id='outer-cell' style='background:gray'>"
            + "<table id='inner' style='display:inline-table;width:100px;table-layout:fixed;background:lime'>"
            + "<colgroup><col span='2' style='width:80px'></colgroup><tr><td></td><td></td></tr></table></td>"
            + "<td id='next-cell' style='width:20px;background:red'></td></tr></table>");

        HtmlRenderShape inner = FixedMinimumShape(rendered, "table#inner");
        HtmlRenderShape outer = FixedMinimumShape(rendered, "td#outer-cell");
        HtmlRenderShape next = FixedMinimumShape(rendered, "td#next-cell");
        Assert.Equal(160D, inner.Width, 3);
        Assert.Equal(inner.Width, outer.Width, 3);
        Assert.True(next.X >= inner.X + inner.Width - 0.001D);
    }

    [Fact]
    public void HtmlTables_FixedFirstRowWidthsSurvivePageHeightContinuationReflow() {
        const string html = "<style>@page{size:300px 30px;margin:0}@page:left{size:300px 40px;margin:0}</style>"
            + "<table id='grid' style='width:100px;table-layout:fixed;background:lime'>"
            + "<tr><td id='r1a' style='width:80px'></td><td id='r1b' style='width:80px'></td></tr>"
            + "<tr><td id='r2a'></td><td id='r2b'></td></tr><tr><td id='r3a'></td><td id='r3b'></td></tr>"
            + "<tr><td id='r4a'></td><td id='r4b'></td></tr></table>";
        HtmlRenderDocument rendered = RenderFixedMinimum(html, paged: true);

        Assert.True(rendered.Pages.Count > 1);
        foreach (string source in new[] { "td#r1a", "td#r1b", "td#r2a", "td#r2b", "td#r3a", "td#r3b", "td#r4a", "td#r4b" }) {
            Assert.Equal(80D, TableIntrinsicGroup(rendered, source).Width, 3);
        }
        HtmlRenderSemanticGroup[] tables = rendered.Pages.SelectMany(page => EnumerateTablePaginationScene(page.Scene))
            .OfType<HtmlRenderSemanticGroup>().Where(group => group.Role == HtmlRenderSemanticGroupRole.Table && group.Source == "table#grid").ToArray();
        Assert.True(tables.Length > 1);
        Assert.All(tables, table => Assert.Equal(160D, table.Width, 3));
    }

    private static HtmlRenderDocument RenderFixedMinimum(string body, bool paged = false) {
        const string defaults = "<!doctype html><style>html,body{margin:0;padding:0;font-size:0;line-height:0}"
            + "table{margin:0;padding:0;border:0;border-spacing:0;border-collapse:separate}"
            + "td{height:12px;padding:0;border:0;vertical-align:top}</style>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(defaults + body, new HtmlRenderOptions {
            Mode = paged ? HtmlRenderMode.Paged : HtmlRenderMode.Continuous,
            UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser,
            ViewportWidth = 300D, ViewportHeight = 100D, Margins = HtmlRenderMargins.All(0D),
            AllowSystemFontFallback = false, FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss
        });
        rendered.RequireNoLoss();
        Assert.Empty(rendered.Text);
        Assert.Empty(rendered.Fonts.Faces);
        return rendered;
    }

    private static HtmlRenderShape FixedMinimumShape(HtmlRenderDocument rendered, string source, OfficeColor? fill = null) =>
        Assert.Single(rendered.Pages.SelectMany(page => FixedMinimumShapes(page.Scene)),
            shape => shape.Source == source && (!fill.HasValue || shape.Shape.FillColor == fill.Value));

    private static IEnumerable<HtmlRenderShape> FixedMinimumShapes(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in EnumerateTablePaginationScene(visuals)) {
            if (visual is HtmlRenderShape shape) yield return shape;
            else if (visual is HtmlRenderClipGroup clip) {
                foreach (HtmlRenderShape child in FixedMinimumShapes(clip.Visuals)) yield return child;
            }
        }
    }
}

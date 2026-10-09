using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    private const string InlineEdgeAtoms = "<span id='a' style='display:inline-block;width:40px;height:12px;background:navy'></span>"
        + "<span id='b' style='display:inline-block;width:50px;height:12px;background:red'></span>";

    // Browser controls use the same two fixed-size boxes, so these contracts do
    // not depend on host typography. Check placement as well as measured width.
    [Theory]
    [InlineData("max-content", "margin-left:-10px", 80D, -10D, 30D, false)]
    [InlineData("max-content", "margin-right:-10px", 80D, 0D, 40D, false)]
    [InlineData("max-content", "margin-inline-start:-10px", 80D, -10D, 30D, false)]
    [InlineData("max-content", "padding-left:10px", 100D, 10D, 50D, false)]
    [InlineData("max-content", "padding-right:10px", 100D, 0D, 40D, false)]
    [InlineData("max-content", "border:2px solid black", 94D, 2D, 42D, false)]
    [InlineData("min-content", "margin-left:-10px", 50D, -10D, 0D, true)]
    [InlineData("min-content", "margin-right:-10px", 40D, 0D, 0D, true)]
    [InlineData("min-content", "padding-right:10px", 60D, 0D, 0D, true)]
    [InlineData("min-content", "padding:0 10px;border:2px solid black", 62D, 12D, 0D, true)]
    [InlineData("min-content", "padding:0 10px;border:2px solid black;box-decoration-break:clone", 74D, 12D, 12D, true)]
    public void HtmlInlineEdges_IntrinsicWidthAndLinePlacementAgree(string width, string edges,
        double expectedWidth, double firstX, double secondX, bool wraps) {
        HtmlRenderDocument rendered = RenderInlineEdgeAtoms(width, edges);
        HtmlRenderShape first = TableGeometryShape(rendered, "span#a");
        HtmlRenderShape second = TableGeometryShape(rendered, "span#b");
        Assert.Equal(expectedWidth, TableGeometryShape(rendered, "div#sized").Width, 3);
        Assert.Equal(firstX, first.X, 3);
        Assert.Equal(secondX, second.X, 3);
        if (wraps) Assert.True(second.Y > first.Y);
        else Assert.Equal(first.Y, second.Y, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("slice", 0D)]
    [InlineData("clone", 12D)]
    public void HtmlInlineEdges_FragmentedBoxRepeatsOnlyClonedEdges(string decorationBreak, double secondX) {
        HtmlRenderDocument rendered = RenderInlineEdgeAtoms("70px", "padding:0 10px;border:2px solid black;box-decoration-break:" + decorationBreak);
        HtmlRenderShape first = TableGeometryShape(rendered, "span#a");
        HtmlRenderShape second = TableGeometryShape(rendered, "span#b");
        Assert.Equal(12D, first.X, 3);
        Assert.Equal(secondX, second.X, 3);
        Assert.True(second.Y > first.Y);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlInlineEdges_OpeningEdgeMovesWithFirstChildWhenPriorContentFillsTheLine() {
        string html = TableGeometrySource("<div style='width:70px'>"
            + "<span id='prior' style='display:inline-block;width:40px;height:12px;background:green'></span>"
            + "<span style='padding-left:10px'><span id='a' style='display:inline-block;width:40px;height:12px;background:navy'></span></span></div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());
        Assert.Equal(10D, TableGeometryShape(rendered, "span#a").X, 3);
        Assert.True(TableGeometryShape(rendered, "span#a").Y > TableGeometryShape(rendered, "span#prior").Y);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlInlineEdges_NestedScopesConsumeEachEdgeOnce() {
        string html = TableGeometrySource("<div id='sized' style='width:max-content;background:lime'>"
            + "<span id='outer' style='padding:0 5px;background:blue'><span id='inner' style='padding:0 6px;background:yellow'>" + InlineEdgeAtoms + "</span></span></div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());
        Assert.Equal(112D, TableGeometryShape(rendered, "div#sized").Width, 3);
        Assert.Equal(11D, TableGeometryShape(rendered, "span#a").X, 3);
        Assert.Equal(51D, TableGeometryShape(rendered, "span#b").X, 3);
        Assert.Equal(0D, TableGeometryShape(rendered, "span#outer").X, 3);
        Assert.Equal(112D, TableGeometryShape(rendered, "span#outer").Width, 3);
        Assert.Equal(5D, TableGeometryShape(rendered, "span#inner").X, 3);
        Assert.Equal(102D, TableGeometryShape(rendered, "span#inner").Width, 3);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlInlineEdges_FloatLineUsesTheSameClosingEdgePreview() {
        string html = TableGeometrySource("<div style='width:110px'><div style='float:left;width:30px;height:40px'></div>"
            + "<span style='margin-right:-10px'>" + InlineEdgeAtoms + "</span></div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());
        Assert.Equal(30D, TableGeometryShape(rendered, "span#a").X, 3);
        Assert.Equal(70D, TableGeometryShape(rendered, "span#b").X, 3);
        Assert.Equal(TableGeometryShape(rendered, "span#a").Y, TableGeometryShape(rendered, "span#b").Y, 3);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlInlineEdges_TopAlignedAtomicChildUsesTheTopOfTheLine() {
        string html = TableGeometrySource("<div id='sized' style='width:100px;background:lime'>"
            + "<span style='padding-left:10px'><span id='a' style='display:inline-block;width:40px;height:12px;vertical-align:top;background:navy'></span></span></div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());
        Assert.Equal(TableGeometryShape(rendered, "div#sized").Y, TableGeometryShape(rendered, "span#a").Y, 3);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlInlineEdges_NoWrapGroupMovesItsClosingEdgeWithItsContent() {
        string html = TableGeometrySource("<div style='width:110px'>"
            + "<span id='prior' style='display:inline-block;width:30px;height:12px;background:green'></span>"
            + "<span style='white-space:nowrap;padding-right:10px'>" + InlineEdgeAtoms + "</span>"
            + "<span id='after' style='display:inline-block;width:15px;height:12px;background:blue'></span></div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());
        HtmlRenderShape first = TableGeometryShape(rendered, "span#a");
        HtmlRenderShape second = TableGeometryShape(rendered, "span#b");
        HtmlRenderShape after = TableGeometryShape(rendered, "span#after");
        Assert.Equal(first.Y, second.Y, 3);
        Assert.True(first.Y > TableGeometryShape(rendered, "span#prior").Y);
        Assert.True(after.Y > second.Y);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlInlineEdges_FirstLineStyleEndsAtTheEdgeAdjustedBreak() {
        string html = TableGeometrySource("<style>#sized::first-line{color:blue}</style>"
            + "<div id='sized' style='width:60px'><span style='padding-left:30px'>AA AA</span></div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());
        HtmlRenderText[] words = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Visuals))
            .OfType<HtmlRenderText>().Where(text => text.Text.Contains("AA", StringComparison.Ordinal)).ToArray();
        Assert.Equal(2, words.Length);
        Assert.Equal(OfficeIMO.Drawing.OfficeColor.Blue, words[0].Color);
        Assert.Equal(OfficeIMO.Drawing.OfficeColor.Black, words[1].Color);
        Assert.True(words[1].Y > words[0].Y);
        rendered.RequireNoLoss();
    }

    private static HtmlRenderDocument RenderInlineEdgeAtoms(string width, string edges) =>
        HtmlRenderTestDriver.Render(TableGeometrySource("<div id='sized' style='width:" + width
            + ";background:lime'><span id='wrapper' style='" + edges + "'>" + InlineEdgeAtoms + "</span></div>"), TableGeometryOptions());
}

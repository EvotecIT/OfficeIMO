using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("slice", 0D, 62D)]
    [InlineData("clone", 12D, 74D)]
    public void HtmlInlineEdges_BlockInterruptionRetainsEnclosingScopeAndRelativeLinkPaint(
        string decorationBreak, double tailStart, double followingStart) {
        string html = TableGeometrySource("<div><a href='https://example.test/' style='position:relative;left:7px;top:5px;"
            + "padding:0 10px;border:2px solid black;background:lime;box-decoration-break:" + decorationBreak + "'>"
            + "<span id='first' style='display:inline-block;width:40px;height:12px;background:navy'></span>"
            + "<span style='display:block;height:40px'>Inside</span>"
            + "<span id='tail' style='display:inline-block;width:50px;height:12px;background:red'></span></a>"
            + "<span id='following' style='display:inline-block;width:15px;height:12px;background:blue'></span></div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());
        HtmlRenderVisual[] visuals = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Visuals)).ToArray();
        HtmlRenderShape Shape(string source) => Assert.Single(visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == source && shape.Shape.FillColor.HasValue);
        HtmlRenderShape first = Shape("span#first");
        HtmlRenderShape tail = Shape("span#tail");
        HtmlRenderShape following = Shape("span#following");
        Assert.Equal(19D, first.X, 3);
        Assert.Equal(7D + tailStart, tail.X, 3);
        Assert.Equal(followingStart, following.X, 3);
        Assert.True(tail.Y > first.Y + 40D);
        Assert.Equal(5D, tail.Y - following.Y, 3);
        Assert.Contains(visuals.OfType<HtmlRenderAnchorFragment>(), anchor => anchor.LinkUri == "https://example.test/"
            && Math.Abs(anchor.Y - tail.Y) < 20D && anchor.X <= tail.X && anchor.X + anchor.Width >= tail.X + tail.Width);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlInlineEdges_IndentAndClosingEdgesShareAtomicFittingWidth() {
        string html = TableGeometrySource("<div style='width:110px;text-indent:30px'>"
            + "<span style='padding:0 10px'>" + InlineEdgeAtoms + "</span></div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());
        HtmlRenderShape first = TableGeometryShape(rendered, "span#a");
        HtmlRenderShape second = TableGeometryShape(rendered, "span#b");
        Assert.Equal(40D, first.X, 3);
        Assert.Equal(0D, second.X, 3);
        Assert.True(second.Y > first.Y);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("overflow-wrap:anywhere", "AAAAAAAA")]
    [InlineData("word-break:break-all", "AAAAAAAA")]
    [InlineData("hyphens:manual", "AA&shy;AAAAAA")]
    public void HtmlInlineEdges_IndentedTokenBreaksRetainClosingAdvance(string declarations, string content) {
        double width = InlineEdgeTextAdvance("AAAAAAAA") + 15D;
        string html = TableGeometrySource("<div style='width:" + width.ToString(System.Globalization.CultureInfo.InvariantCulture)
            + "px;text-indent:20px'><span style='padding-left:10px;padding-right:10px;" + declarations + "'>"
            + content + "</span></div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());
        HtmlRenderText[] texts = InlineEdgeTexts(rendered);
        Assert.Equal("AAAAAAAA", string.Concat(texts.Select(text => text.Text)).Replace("-", string.Empty));
        Assert.Equal(30D, texts[0].X, 3);
        Assert.True(texts.Select(text => text.Y).Distinct().Count() > 1);
        Assert.All(texts, text => Assert.True(text.X + (text.TextAdvanceWidth ?? text.Width) <= width + 0.001D));
        Assert.True(texts[texts.Length - 1].X + (texts[texts.Length - 1].TextAdvanceWidth ?? texts[texts.Length - 1].Width)
            + 10D <= width + 0.001D);
        rendered.RequireNoLoss();
    }
}

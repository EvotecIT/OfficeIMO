using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlInlineEdges_TrimmingOutsideWhitespaceKeepsOneClosingEdge() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(TableGeometrySource(
            "<div style='width:100px;text-align:right'><span id='outer' style='padding-right:10px;background:blue'>AA</span> </div>"), TableGeometryOptions());
        HtmlRenderText text = Assert.Single(InlineEdgeTexts(rendered));
        double advance = text.TextAdvanceWidth ?? text.Width;
        Assert.Equal(100D - advance - 10D, text.X, 3);
        Assert.Equal(advance + 10D, TableGeometryShape(rendered, "span#outer").Width, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlInlineEdges_EmergencyBreakIncludesThePendingOpeningEdge(bool floated) {
        double wordWidth = InlineEdgeTextAdvance("AAAAAA");
        double width = wordWidth + 10D + (floated ? 30D : 0D);
        string content = "<div style='width:" + width.ToString(System.Globalization.CultureInfo.InvariantCulture) + "px'>"
            + (floated ? "<div style='float:left;width:30px;height:100px'></div>" : string.Empty)
            + "<span style='padding-left:20px;overflow-wrap:anywhere'>AAAAAA</span></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(TableGeometrySource(content), TableGeometryOptions());
        HtmlRenderText[] texts = InlineEdgeTexts(rendered);
        Assert.Equal("AAAAAA", string.Concat(texts.Select(text => text.Text)));
        Assert.True(texts.Select(text => text.Y).Distinct().Count() > 1);
        Assert.Equal(20D + (floated ? 30D : 0D), texts[0].X, 3);
        Assert.All(texts, text => Assert.True(text.X + (text.TextAdvanceWidth ?? text.Width) <= width + 0.001D));
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlInlineEdges_FloatManualHyphenationReservesTheWholeRemainderClosingEdge() {
        double wordWidth = InlineEdgeTextAdvance("AAAAAA");
        string content = "<div style='width:" + (wordWidth + 40D).ToString(System.Globalization.CultureInfo.InvariantCulture)
            + "px'><div style='float:left;width:30px;height:100px'></div>"
            + "<span style='padding-right:20px;hyphens:manual'>AA&shy;AAAA</span></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(TableGeometrySource(content), TableGeometryOptions());
        HtmlRenderText[] texts = InlineEdgeTexts(rendered);
        Assert.Equal(new[] { "AA-", "AAAA" }, texts.Select(text => text.Text).ToArray());
        Assert.True(texts[1].Y > texts[0].Y);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlInlineEdges_FloatObstructionTestsTheAtomicBoxTogetherWithItsEdges() {
        string content = "<div style='width:100px'><div style='float:left;width:30px;height:40px'></div>"
            + "<span style='padding-left:20px'><span id='a' style='display:inline-block;width:60px;height:12px;background:navy'></span></span></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(TableGeometrySource(content), TableGeometryOptions());
        HtmlRenderShape atom = TableGeometryShape(rendered, "span#a");
        Assert.Equal(20D, atom.X, 3);
        Assert.True(atom.Y >= 40D);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlInlineEdges_FirstLineTokenSplitIncludesTheClosingEdge() {
        double wordWidth = InlineEdgeTextAdvance("AAAAAA");
        string content = "<style>#sized::first-line{color:blue}</style><div id='sized' style='width:"
            + (wordWidth + 10D).ToString(System.Globalization.CultureInfo.InvariantCulture)
            + "px'><span style='padding-right:20px;overflow-wrap:anywhere'>AAAAAA</span></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(TableGeometrySource(content), TableGeometryOptions());
        HtmlRenderText[] texts = InlineEdgeTexts(rendered);
        Assert.Equal("AAAAAA", string.Concat(texts.Select(text => text.Text)));
        Assert.Equal(OfficeIMO.Drawing.OfficeColor.Blue, texts[0].Color);
        Assert.Equal(OfficeIMO.Drawing.OfficeColor.Black, texts[texts.Length - 1].Color);
        Assert.True(texts[texts.Length - 1].Y > texts[0].Y);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlInlineEdges_EmergencyFragmentsReserveOnlyTheFinalSlicedClosingEdge(bool floated) {
        double available = InlineEdgeTextAdvance("A") * 4D + 0.01D;
        string content = "<div style='width:" + (available + (floated ? 30D : 0D)).ToString(System.Globalization.CultureInfo.InvariantCulture)
            + "px'>" + (floated ? "<div style='float:left;width:30px;height:100px'></div>" : string.Empty)
            + "<span style='padding-right:10px;overflow-wrap:anywhere'>AAAAAAAA</span></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(TableGeometrySource(content), TableGeometryOptions());
        HtmlRenderText[] texts = InlineEdgeTexts(rendered);
        Assert.Equal("AAAA", texts[0].Text);
        Assert.Equal("AAAAAAAA", string.Concat(texts.Select(text => text.Text)));
        Assert.True(texts.Length >= 2);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("slice", false, false)]
    [InlineData("clone", false, false)]
    [InlineData("clone", true, false)]
    [InlineData("slice", false, true)]
    public void HtmlInlineEdges_PreservedTabsUseTheContentCursorBeforeReservedCloneEnds(string decorationBreak, bool floated, bool firstLine) {
        double stop = (InlineEdgeTextAdvance("A A") - 2D * InlineEdgeTextAdvance("A")) * 4D;
        double expectedX = (Math.Floor((10D + InlineEdgeTextAdvance("A")) / stop) + 1D) * stop;
        string content = (firstLine ? "<style>#sized::first-line{color:blue}</style>" : string.Empty)
            + "<div id='sized'>" + (floated ? "<div style='float:left;width:30px;height:100px'></div>" : string.Empty)
            + "<span style='padding:0 10px;white-space:pre;tab-size:4;box-decoration-break:" + decorationBreak + "'>A\tX</span></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(TableGeometrySource(content), TableGeometryOptions());
        HtmlRenderText x = Assert.Single(InlineEdgeTexts(rendered), text => text.Text == "X");
        Assert.Equal(expectedX + (floated ? 30D : 0D), x.X, 3);
        if (firstLine) Assert.Equal(OfficeIMO.Drawing.OfficeColor.Blue, x.Color);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlInlineEdges_MovingANoWrapSuffixKeepsTheSlicedClosingEdgeOnItsLastLine(bool floated) {
        string content = "<div style='width:" + (floated ? "100" : "70") + "px;text-align:right'>"
            + (floated ? "<div style='float:left;width:30px;height:80px'></div>" : string.Empty)
            + "<span style='padding-right:10px'>AA <span style='white-space:nowrap'>AA AA</span></span></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(TableGeometrySource(content), TableGeometryOptions());
        HtmlRenderText[] texts = InlineEdgeTexts(rendered);
        Assert.Equal(new[] { "AA", "AA AA" }, texts.Select(text => text.Text).ToArray());
        double origin = floated ? 30D : 0D;
        Assert.Equal(origin + 70D - (texts[0].TextAdvanceWidth ?? texts[0].Width), texts[0].X, 3);
        Assert.Equal(origin + 70D - (texts[1].TextAdvanceWidth ?? texts[1].Width) - 10D, texts[1].X, 3);
        Assert.True(texts[1].Y > texts[0].Y);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlInlineEdges_AncestorDecorationAndLinkUseTheirOwnRelativePaintOffset() {
        string content = "<div id='sized' style='width:max-content;background:lime'>"
            + "<a id='outer' href='https://example.com' style='padding:0 5px;background:blue'>"
            + "<span style='position:relative;left:10px'>AAAA</span><span>BBBB</span></a></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(TableGeometrySource(content), TableGeometryOptions());
        double width = TableGeometryShape(rendered, "div#sized").Width;
        HtmlRenderShape outer = TableGeometryShape(rendered, "a#outer");
        HtmlRenderAnchorFragment anchor = Assert.Single(rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Visuals)).OfType<HtmlRenderAnchorFragment>());
        Assert.Equal(0D, outer.X, 3);
        Assert.Equal(width, outer.Width, 3);
        Assert.Equal(0D, anchor.X, 3);
        Assert.Equal(width, anchor.Width, 3);
        Assert.Equal(15D, Assert.Single(InlineEdgeTexts(rendered), text => text.Text == "AAAA").X, 3);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlInlineEdges_EmptyDecoratedDescendantsRespectTheLayoutOperationBudget() {
        string content = "<div>" + string.Concat(Enumerable.Repeat("<span style='padding-left:1px'></span>", 64)) + "</div>";
        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlRenderTestDriver.Render(content, new HtmlRenderOptions { MaxLayoutOperations = 8 }));
        Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutOperations), exception.LimitSource);
    }

    private static HtmlRenderText[] InlineEdgeTexts(HtmlRenderDocument rendered) =>
        rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Visuals)).OfType<HtmlRenderText>().ToArray();

    private static double InlineEdgeTextAdvance(string text) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(TableGeometrySource("<span>" + text + "</span>"), TableGeometryOptions());
        return InlineEdgeTexts(rendered).Sum(item => item.TextAdvanceWidth ?? item.Width);
    }
}

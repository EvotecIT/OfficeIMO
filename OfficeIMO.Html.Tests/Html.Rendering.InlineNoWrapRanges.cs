using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false, "slice")]
    [InlineData(true, "slice")]
    [InlineData(false, "clone")]
    [InlineData(true, "clone")]
    public void HtmlInlineEdges_MultiRunNoWrapRangeMovesNestedEdgesTogether(bool floated, string decorationBreak) {
        string content = string.Concat(Enumerable.Range(0, 8).Select(index =>
            "<span style='color:" + (index % 2 == 0 ? "navy" : "red") + "'>AA</span>"));
        string html = TableGeometrySource("<div style='width:" + (floated ? "330" : "300") + "px;text-align:right'>"
            + (floated ? "<div style='float:left;width:30px;height:100px'></div>" : string.Empty)
            + "<span style='padding-left:10px;padding-right:20px;box-decoration-break:" + decorationBreak + "'>"
            + "<span id='prior' style='display:inline-block;width:100px;height:12px;background:lime'></span>"
            + "<span style='white-space:nowrap;padding-left:3px;padding-right:4px'>" + content + "</span></span></div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());
        HtmlRenderText[] text = InlineEdgeTexts(rendered);
        Assert.Equal(8, text.Length);
        Assert.All(text, item => Assert.Equal("AA", item.Text));
        Assert.Single(text.Select(item => item.Y).Distinct());
        Assert.True(text[0].Y > TableGeometryShape(rendered, "span#prior").Y);
        Assert.Equal((floated ? 330D : 300D) - 4D - 20D,
            text[text.Length - 1].X + (text[text.Length - 1].TextAdvanceWidth ?? text[text.Length - 1].Width), 3);
        for (int index = 1; index < text.Length; index++) {
            Assert.Equal(text[index - 1].X + (text[index - 1].TextAdvanceWidth ?? text[index - 1].Width), text[index].X, 3);
        }
        HtmlRenderTestDriver.AssertOnlyUnselectedFontEstimates(rendered);
    }
}

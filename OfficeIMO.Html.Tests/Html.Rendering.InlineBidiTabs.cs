using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("slice")]
    [InlineData("clone")]
    public void HtmlInlineEdges_IdentityBidiOverridePreservesTabStopsAndSourceEdges(string decorationBreak) {
        HtmlRenderDocument plain = RenderInlineBidiTab("A\tX", decorationBreak);
        HtmlRenderDocument overridden = RenderInlineBidiTab("\u202dA\tX\u202c", decorationBreak);
        foreach (string value in new[] { "A", "X", "END" }) {
            HtmlRenderText expected = Assert.Single(InlineEdgeTexts(plain), text => text.Text == value);
            HtmlRenderText actual = Assert.Single(InlineEdgeTexts(overridden), text => text.Text == value);
            Assert.Equal(expected.X, actual.X, 3);
            Assert.Equal(expected.Y, actual.Y, 3);
        }
    }

    [Theory]
    [InlineData("slice")]
    [InlineData("clone")]
    public void HtmlInlineEdges_ReversedTabSegmentsKeepAdvancesOnTheirOwnSource(string decorationBreak) {
        HtmlRenderDocument rendered = RenderInlineBidiTab("\u202eA\tX\u202c", decorationBreak);
        HtmlRenderText[] text = InlineEdgeTexts(rendered);
        HtmlRenderText first = Assert.Single(text, item => item.Text == "X");
        HtmlRenderText last = Assert.Single(text, item => item.Text == "A");
        HtmlRenderText after = Assert.Single(text, item => item.Text == "END");
        HtmlRenderText plainAfter = Assert.Single(InlineEdgeTexts(RenderInlineBidiTab("A\tX", decorationBreak)), item => item.Text == "END");

        // The original A segment owns the opening edge; X owns the closing edge.
        // Reordering their shared run must not transfer either edge to the other segment.
        Assert.Equal(0D, first.X, 3);
        Assert.True(last.X > first.X + (first.TextAdvanceWidth ?? first.Width) + 20D);
        Assert.Equal(plainAfter.X, after.X, 3);
    }

    private static HtmlRenderDocument RenderInlineBidiTab(string text, string decorationBreak) =>
        HtmlRenderTestDriver.Render(TableGeometrySource("<div style='width:400px'>"
            + "<span style='padding-left:10px;padding-right:20px;white-space:pre;tab-size:4;box-decoration-break:"
            + decorationBreak + "'>" + text + "</span><span>END</span></div>"), TableGeometryOptions());
}

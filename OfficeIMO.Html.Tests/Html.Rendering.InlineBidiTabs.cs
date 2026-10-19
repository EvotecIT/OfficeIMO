using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("slice", "\u202d")]
    [InlineData("clone", "\u202d")]
    [InlineData("slice", "\u202e")]
    [InlineData("clone", "\u202e")]
    public void HtmlInlineEdges_BidiOverridePreservesTabStopsAndSourceEdges(string decorationBreak, string overrideControl) {
        // UAX #9 L1 resets a tab to the paragraph level, so even a right-to-left override keeps each
        // single-character segment on its own side of the tab stop with its own advance and edges.
        HtmlRenderDocument plain = RenderInlineBidiTab("A\tX", decorationBreak);
        HtmlRenderDocument overridden = RenderInlineBidiTab(overrideControl + "A\tX\u202c", decorationBreak);
        foreach (string value in new[] { "A", "X", "END" }) {
            HtmlRenderText expected = Assert.Single(InlineEdgeTexts(plain), text => text.Text == value);
            HtmlRenderText actual = Assert.Single(InlineEdgeTexts(overridden), text => text.Text == value);
            Assert.Equal(expected.X, actual.X, 3);
            Assert.Equal(expected.Y, actual.Y, 3);
        }
    }

    private static HtmlRenderDocument RenderInlineBidiTab(string text, string decorationBreak) =>
        HtmlRenderTestDriver.Render(TableGeometrySource("<div style='width:400px'>"
            + "<span style='padding-left:10px;padding-right:20px;white-space:pre;tab-size:4;box-decoration-break:"
            + decorationBreak + "'>" + text + "</span><span>END</span></div>"), TableGeometryOptions());
}

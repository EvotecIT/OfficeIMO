using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("inline-block")]
    [InlineData("inline-flex")]
    [InlineData("inline-grid")]
    public void HtmlRender_FlexWrapperPreservesTheNaturalWidthOfADecoratedAtomicChild(string display) {
        string css = "<style>body{margin:0;font:12px Arial}#chip{display:" + display
            + ";padding:2px 6px;border:1px solid blue;white-space:nowrap}</style>";
        var options = new HtmlRenderOptions { ViewportWidth=400D, Margins=HtmlRenderMargins.All(0D) };
        HtmlRenderDocument standalone = HtmlRenderTestDriver.Render(css+"<div><span id='chip'>Alpha beta</span></div>",options);
        HtmlRenderDocument wrapped = HtmlRenderTestDriver.Render(css
            + "<div style='display:flex;gap:4px'><div><span id='chip'>Alpha beta</span></div><div id='following'>Next</div></div>",options);
        HtmlRenderShape expected = Assert.Single(standalone.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source=="span#chip" && shape.Shape.StrokeWidth>0D);
        HtmlRenderShape actual = Assert.Single(wrapped.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source=="span#chip" && shape.Shape.StrokeWidth>0D);
        Assert.Equal(expected.Width,actual.Width,3);
        Assert.Equal(expected.Height,actual.Height,3);
        HtmlRenderText following = Assert.Single(wrapped.Pages[0].Visuals.OfType<HtmlRenderText>(),text=>text.Text=="Next");
        Assert.True(following.X>=actual.X+actual.Width+4D-.001D,
            "The following flex item must reserve the atomic child's border and padding rather than its text width alone.");
    }
}

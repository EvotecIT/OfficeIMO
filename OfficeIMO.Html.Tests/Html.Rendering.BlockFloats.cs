using OfficeIMO.Html;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlBlockFloatTests {
    [Theory]
    [InlineData("margin-left:40px", "float:left;width:40px", 40D)]
    [InlineData("margin-left:40px", "float:left;width:40px;margin-left:-40px", 40D)]
    [InlineData("", "float:left;width:40px", 40D)]
    public void FloatAndFollowingBlockShareFirstLine(string body, string label, double expectedX) {
        var text = Text("<span style='" + label + "'>Label</span><div style='" + body + "'>Body</div>");
        var number = Assert.Single(text, t => t.Text == "Label");
        var main = Assert.Single(text, t => t.Text == "Body");
        Assert.Equal(number.Y, main.Y, 3);
        Assert.Equal(expectedX, main.X, 3);
    }

    [Theory]
    [InlineData("left", 60D)]
    [InlineData("right", 40D)]
    [InlineData("both", 60D)]
    public void ClearBlockAdvancesPastRequestedFloatSide(string clear, double expectedY) {
        var text = Text("<div style='float:left;width:40px;height:60px'>Left</div><div style='float:right;width:40px;height:40px'>Right</div><div style='clear:" + clear + "'>Body</div>");
        Assert.Equal(expectedY, Assert.Single(text, t => t.Text == "Body").Y, 3);
    }

    [Fact]
    public void FloatInfluenceContinuesAcrossNestedAndFollowingBlocks() {
        var text = Text("<div style='float:left;width:40px;height:40px'>Label</div><section><div>First</div></section><div>Second</div><div>Third</div>");
        Assert.Equal(40D, Assert.Single(text, t => t.Text == "First").X, 3);
        Assert.Equal(40D, Assert.Single(text, t => t.Text == "Second").X, 3);
        Assert.Equal(0D, Assert.Single(text, t => t.Text == "Third").X, 3);
        Assert.Equal(40D, Assert.Single(text, t => t.Text == "Third").Y, 3);
    }

    [Theory]
    [InlineData("display:flow-root")]
    [InlineData("overflow:hidden")]
    public void IndependentBlockFitsBesideExternalFloat(string context) {
        var text = Text("<div style='float:left;width:40px;height:60px'>Label</div><div style='" + context + ";width:100px'>Body</div>");
        Assert.Equal(40D, Assert.Single(text, t => t.Text == "Body").X, 3);
        Assert.Equal(0D, Assert.Single(text, t => t.Text == "Body").Y, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FollowingBlockBackgroundDoesNotCoverTheFloat(bool paged) {
        var rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
            "<style>@page{size:300px 100px;margin:0}</style><div><div style='float:left;width:40px;height:40px;background:red'></div><div style='height:40px;background:blue'>Body</div></div>"),
            new HtmlRenderOptions { Mode = paged ? HtmlRenderMode.Paged : HtmlRenderMode.Continuous, ViewportWidth = 300, Margins = HtmlRenderMargins.All(0) });
        var raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        Assert.Equal(255, raster.GetPixel(10, 10).R);
        Assert.Equal(255, raster.GetPixel(100, 30).B);
    }

    [Fact]
    public void WideIndependentBlockMovesBelowTheFloat() {
        var text = Text("<div style='float:left;width:40px;height:60px'>Label</div><div style='display:flow-root;width:280px'>Body</div>");
        var body = Assert.Single(text, t => t.Text == "Body");
        Assert.Equal(0D, body.X, 3);
        Assert.Equal(60D, body.Y, 3);
    }

    [Fact]
    public void CollapsedSiblingMarginsUseTheActualFloatBand() {
        var text = Text("<div style='float:left;width:40px;height:60px'>Label</div><div style='margin-bottom:20px'>First</div><div style='margin-top:20px'>Second</div>");
        var second = Assert.Single(text, t => t.Text == "Second");
        Assert.Equal(40D, second.Y, 3);
        Assert.Equal(40D, second.X, 3);
    }

    [Fact]
    public void FloatPaintAndTextSurvivePageFragmentsWithoutRepeatingTheLabel() {
        var rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
            "<style>@page{size:300px 100px;margin:0}</style><div style='font:10px Arial;line-height:20px'><div style='float:left;width:40px;height:140px;background:red'>Label</div><div style='height:140px;background:blue'>Body</div></div>"),
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, Margins = HtmlRenderMargins.All(0) });
        Assert.Equal(2, rendered.Pages.Count);
        Assert.Single(rendered.Pages.SelectMany(p => Flatten(p.Scene)).OfType<HtmlRenderText>(), t => t.Text == "Label");
        foreach (var page in rendered.Pages) {
            var raster = OfficeDrawingRasterRenderer.Render(page.CreateDrawing());
            Assert.Equal(255, raster.GetPixel(10, 30).R);
            Assert.Equal(255, raster.GetPixel(100, 30).B);
        }
    }

    [Fact]
    public void OutsideListMarkerDoesNotShiftTheFloatExclusionTwice() {
        var text = Text("<div style='float:left;width:40px;height:40px'>Label</div><ul style='margin:0;padding:0;list-style-position:outside'><li>Body words beside a float</li></ul>");
        var body = Assert.Single(text, t => t.Text.StartsWith("Body", StringComparison.Ordinal));
        Assert.Equal(40D, body.X, 3);
    }

    private static HtmlRenderText[] Text(string body) => HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
        "<div style='font:10px Arial;line-height:20px'>" + body + "</div>"),
        new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 300, Margins = HtmlRenderMargins.All(0) })
        .Pages.SelectMany(p => Flatten(p.Scene)).OfType<HtmlRenderText>().ToArray();

    private static IEnumerable<HtmlRenderVisual> Flatten(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (var visual in visuals) {
            yield return visual;
            IEnumerable<HtmlRenderVisual>? children = visual switch {
                HtmlRenderClipGroup g => g.Visuals,
                HtmlRenderPathClipGroup g => g.Visuals,
                HtmlRenderEffectGroup g => g.Visuals,
                HtmlRenderSemanticGroup g => g.Visuals,
                HtmlRenderLogicalTextGroup g => g.Visuals,
                _ => null
            };
            if (children != null) foreach (var child in Flatten(children)) yield return child;
        }
    }

}

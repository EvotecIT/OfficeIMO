using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("", 104D, 54D)]
    [InlineData("box-sizing:content-box", 104D, 54D)]
    [InlineData("box-sizing:border-box", 80D, 40D)]
    [InlineData("--box:border-box;box-sizing:var(--box)", 80D, 40D)]
    [InlineData("box-sizing:border-box!important;box-sizing:content-box", 80D, 40D)]
    [InlineData("box-sizing:border-box;box-sizing:initial", 104D, 54D)]
    [InlineData("box-sizing:border-box;box-sizing:unset", 104D, 54D)]
    public void HtmlPageMarginBoxSizing_AuthoredDimensionsUseTheSelectedBoxEdges(
        string declaration, double expectedWidth, double expectedHeight) {
        HtmlRenderDocument rendered = RenderPageMarginDimensions(
            "width:80px;height:40px;padding:5px 10px;border:2px solid red;" + declaration, "''");
        HtmlRenderPage page = Assert.Single(rendered.Pages);

        AssertPageMarginDimensionFrames(rendered, 60D - expectedHeight, 240D, expectedHeight);
        foreach (string position in new[] { "top-center", "bottom-center" }) {
            HtmlRenderSemanticGroup box = PageMarginGeometryBox(page, position);
            Assert.Equal(expectedWidth, box.Width, 6);
            Assert.Equal(170D, box.X + box.Width / 2D, 6);
            Assert.Equal(4, box.Visuals.OfType<HtmlRenderShape>().Count());
        }
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("width:50%;height:50%", 150D, 30D)]
    [InlineData("width:100px;max-width:80px;height:50px;max-height:40px", 80D, 40D)]
    [InlineData("width:40px;min-width:100px;max-width:80px;height:20px;min-height:50px;max-height:40px", 100D, 50D)]
    [InlineData("max-width:80px;max-height:40px", 80D, 40D)]
    [InlineData("width:0;height:0", 24D, 14D)]
    public void HtmlPageMarginBoxSizing_NumericConstraintsKeepInsetFloorsAndMinimumPrecedence(
        string declaration, double expectedWidth, double expectedHeight) {
        HtmlRenderDocument rendered = RenderPageMarginDimensions(
            "box-sizing:border-box;padding:5px 10px;border:2px solid red;" + declaration, "''");
        HtmlRenderPage page = Assert.Single(rendered.Pages);

        AssertPageMarginDimensionFrames(rendered, 60D - expectedHeight, 240D, expectedHeight);
        Assert.Equal(expectedWidth, PageMarginGeometryBox(page, "top-center").Width, 6);
        Assert.Equal(expectedWidth, PageMarginGeometryBox(page, "bottom-center").Width, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("min-content", "", 50D, 2)]
    [InlineData("max-content", "", 100D, 1)]
    [InlineData("fit-content", "", 100D, 1)]
    [InlineData("max-content", "max-width:84px", 60D, 2)]
    [InlineData("min-content", "min-width:124px", 100D, 1)]
    public void HtmlPageMarginBoxSizing_IntrinsicWidthsStayContentBasedBesideNumericLimits(
        string keyword, string constraint, double expectedContentWidth, int expectedLines) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:'ALPHA BETA';box-sizing:border-box;width:" + keyword
            + ";padding:0 10px;border:2px solid red;" + constraint + "}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(page, "top-center");
        HtmlRenderText text = PageMarginTypographyText(page, "top-center");

        Assert.Equal(expectedContentWidth + 24D, box.Width, 6);
        Assert.Equal(expectedContentWidth, text.Width, 6);
        Assert.Equal(box.X + 12D, text.X, 6);
        Assert.Equal(expectedLines, MarginSizingPaintedText(page, text).Length);
        Assert.Equal("ALPHA BETA", text.Text);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("", 104D, 54D)]
    [InlineData("box-sizing:inherit", 80D, 40D)]
    public void HtmlPageMarginBoxSizing_PageContextRequiresExplicitInheritance(
        string declaration, double expectedWidth, double expectedHeight) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "box-sizing:border-box;@top-center{content:'';width:80px;height:40px;"
            + "padding:5px 10px;border:2px solid red;" + declaration + "}");
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(Assert.Single(rendered.Pages), "top-center");

        Assert.Equal(expectedWidth, box.Width, 6);
        Assert.Equal(expectedHeight, box.Height, 6);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginBoxSizing_SignedMarginsAndAutomaticPeerKeepTheirOuterAnchors() {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-left{content:'';box-sizing:border-box;width:80px;height:40px;"
            + "padding:5px 10px;border:2px solid red;margin:0 20px 0 -10px}"
            + "@top-right{content:'R'}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup left = PageMarginGeometryBox(page, "top-left");
        HtmlRenderText right = PageMarginTypographyText(page, "top-right");

        Assert.Equal(10D, left.X, 6);
        Assert.Equal(80D, left.Width, 6);
        Assert.Equal(110D, right.X, 6);
        Assert.Equal(210D, right.Width, 6);
        Assert.Equal(320D, right.X + right.Width, 6);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginBoxSizing_RunningContentUsesTheBorderBoxInnerFrame() {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:element(heading);box-sizing:border-box;"
            + "width:80px;height:40px;padding:5px 10px;border:2px solid red}",
            body: "<p style='position:running(heading);height:20px;background:blue'>A</p><p>BODY</p>");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup box = Assert.Single(page.Scene.OfType<HtmlRenderSemanticGroup>(),
            group => group.Source == "@page @top-center element(heading)");
        HtmlRenderShape background = Assert.Single(EnumerateRenderVisuals(box.Visuals).OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillColor == OfficeColor.Blue);

        Assert.Equal(80D, box.Width, 6);
        Assert.Equal(40D, box.Height, 6);
        Assert.Equal(box.X + 12D, background.X, 6);
        Assert.Equal(56D, background.Width, 6);
        Assert.Equal(20D, background.Height, 6);
        Assert.Contains(PageMarginTypographyDrawingText(page.CreateDrawing()), text => text.Text == "A");
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginBoxSizing_RunningPaintOutsideTheReducedContentHeightReportsLoss() {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:element(heading);box-sizing:border-box;"
            + "width:80px;height:30px;padding:5px 10px;border:2px solid red}",
            body: "<p style='position:running(heading);height:20px;background:blue'>A</p><p>BODY</p>");

        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Source == "@page @top-center element(heading)"
            && diagnostic.Detail == "margin-height" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.ThrowsAny<InvalidOperationException>(() => rendered.RequireNoLoss());
    }
}

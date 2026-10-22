using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("top", "min-content", 50D, 2)]
    [InlineData("bottom", "min-content", 50D, 2)]
    [InlineData("top", "max-content", 100D, 1)]
    [InlineData("bottom", "max-content", 100D, 1)]
    public void HtmlPageMarginIntrinsicWidths_CentredTextUsesItsAuthoredContentWidth(
        string side, string keyword, double expectedWidth, int expectedLines) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@" + side + "-center{content:'ALPHA BETA';width:" + keyword + "}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(page, side + "-center");
        HtmlRenderText text = PageMarginTypographyText(page, side + "-center");

        Assert.Equal(expectedWidth, box.Width, 6);
        Assert.Equal(170D, box.X + box.Width / 2D, 6);
        Assert.Equal(side == "top" ? 0D : 240D, box.Y, 6);
        Assert.Equal(60D, box.Height, 6);
        Assert.Equal(expectedWidth, text.Width, 6);
        Assert.Equal(expectedLines, MarginSizingPaintedText(page, text).Length);
        Assert.Equal("ALPHA BETA", text.Text);
        Assert.Equal(HtmlRenderSemanticGroupRole.Artifact, box.Role);
        Assert.Equal("BODY", rendered.Text.Trim());
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("width:120px;max-width:min-content", 50D)]
    [InlineData("width:20px;min-width:max-content", 100D)]
    [InlineData("width:120px;min-width:max-content;max-width:min-content", 100D)]
    [InlineData("width:20px;min-width:fit-content", 100D)]
    [InlineData("width:120px;max-width:fit-content", 100D)]
    public void HtmlPageMarginIntrinsicWidths_KeywordConstraintsOverrideThePreferredNumericWidth(
        string declaration, double expectedWidth) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:'ALPHA BETA';" + declaration + "}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderText text = PageMarginTypographyText(page, "top-center");

        Assert.Equal(expectedWidth, text.Width, 6);
        Assert.Equal(170D, text.X + text.Width / 2D, 6);
        Assert.Equal("ALPHA BETA", text.Text);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("ALPHA BETA", 100D, 1)]
    [InlineData("ALPHA BETA GAMMA DELTA EPSILON ZETA", 296D, 2)]
    public void HtmlPageMarginIntrinsicWidths_FitContentUsesTheSpaceInsideItsBorder(
        string content, double expectedContentWidth, int expectedLines) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:'" + content + "';width:fit-content;border:2px solid red}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(page, "top-center");
        HtmlRenderText text = PageMarginTypographyText(page, "top-center");

        Assert.Equal(expectedContentWidth + 4D, box.Width, 6);
        Assert.Equal(170D, box.X + box.Width / 2D, 6);
        Assert.Equal(box.X + 2D, text.X, 6);
        Assert.Equal(expectedContentWidth, text.Width, 6);
        Assert.Equal(expectedLines, MarginSizingPaintedText(page, text).Length);
        Assert.Equal(content, text.Text);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("min-content", 50D)]
    [InlineData("max-content", 100D)]
    public void HtmlPageMarginIntrinsicWidths_AutoSiblingKeepsTheOppositeOuterAnchor(
        string keyword, double expectedLeftWidth) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-left{content:'ALPHA BETA';width:" + keyword + "}@top-right{content:'R'}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderText left = PageMarginTypographyText(page, "top-left");
        HtmlRenderText right = PageMarginTypographyText(page, "top-right");

        Assert.Equal(20D, left.X, 6);
        Assert.Equal(expectedLeftWidth, left.Width, 6);
        Assert.Equal(300D - expectedLeftWidth, right.Width, 6);
        Assert.Equal(left.X + left.Width, right.X, 6);
        Assert.Equal(320D, right.X + right.Width, 6);
        Assert.Equal("ALPHA BETA", left.Text);
        Assert.Equal("R", right.Text);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("min-content", 50D)]
    [InlineData("max-content", 100D)]
    public void HtmlPageMarginIntrinsicWidths_PaddingAndBorderSurroundTheMeasuredContentOnce(
        string keyword, double expectedContentWidth) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:'ALPHA BETA';width:" + keyword + ";padding:0 10px;border:2px solid red}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(page, "top-center");
        HtmlRenderText text = PageMarginTypographyText(page, "top-center");

        Assert.Equal(expectedContentWidth + 24D, box.Width, 6);
        Assert.Equal(170D, box.X + box.Width / 2D, 6);
        Assert.Equal(box.X + 12D, text.X, 6);
        Assert.Equal(expectedContentWidth, text.Width, 6);
        Assert.Contains(box.Visuals.OfType<HtmlRenderShape>(), shape => shape.Shape.StrokeColor == OfficeColor.Red);
        Assert.Equal("ALPHA BETA", text.Text);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginIntrinsicWidths_EmptyGeneratedBoxRetainsOnlyItsAuthoredInsets() {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:'';width:max-content;padding:0 6px;border:2px solid red}"
            + "@bottom-center{content:none;width:max-content;padding:0 6px;border:2px solid red}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(page, "top-center");

        Assert.Equal(16D, box.Width, 6);
        Assert.Equal(162D, box.X, 6);
        Assert.Contains(box.Visuals.OfType<HtmlRenderShape>(), shape => shape.Shape.StrokeColor == OfficeColor.Red);
        Assert.Empty(EnumerateRenderVisuals(box.Visuals).OfType<HtmlRenderText>());
        Assert.DoesNotContain(page.Scene.OfType<HtmlRenderSemanticGroup>(), group => group.Source == "@page @bottom-center");
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("max-content", "'A\\A BBBB'", 40D, "A|BBBB")]
    [InlineData("min-content", "'UNBREAKABLE'", 110D, "UNBREAKABLE")]
    public void HtmlPageMarginIntrinsicWidths_HardLinesAndUnbreakableContentKeepTheirMeasuredPaint(
        string keyword, string content, double expectedWidth, string expectedLines) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:" + content + ";width:" + keyword + "}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderText text = PageMarginTypographyText(page, "top-center");

        Assert.Equal(expectedWidth, text.Width, 6);
        Assert.Equal(expectedLines.Split('|'), MarginSizingPaintedText(page, text).Select(line => line.Text).ToArray());
        Assert.Equal(PageMarginTypographyCompact(expectedLines.Replace("|", "")), PageMarginTypographyCompact(text.Text));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("min-content", 50D)]
    [InlineData("max-content", 100D)]
    public void HtmlPageMarginIntrinsicWidths_RunningContentUsesItsAssignedInnerWidth(
        string keyword, double expectedContentWidth) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:element(heading);width:" + keyword + ";padding:0 10px;border:2px solid red}",
            body: "<p id='running' style='position:running(heading)'>ALPHA BETA</p><p>BODY</p>");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup box = Assert.Single(page.Scene.OfType<HtmlRenderSemanticGroup>(),
            group => group.Source == "@page @top-center element(heading)");
        HtmlRenderText[] lines = EnumerateRenderVisuals(box.Visuals).OfType<HtmlRenderText>().ToArray();

        Assert.Equal(expectedContentWidth + 24D, box.Width, 6);
        Assert.Equal(170D, box.X + box.Width / 2D, 6);
        Assert.Equal(keyword == "min-content" ? 2 : 1, lines.Length);
        Assert.All(lines, text => {
            Assert.Equal(box.X + 12D, text.X, 6);
            Assert.InRange(text.Width, 0.01D, expectedContentWidth);
            Assert.InRange(text.X + text.Width, box.X + 12D, box.X + 12D + expectedContentWidth);
        });
        Assert.Equal("ALPHABETA", PageMarginTypographyCompact(string.Concat(lines.Select(text => text.Text))));
        Assert.Equal("BODY", rendered.Text.Trim());
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginIntrinsicWidths_CountersUseTheirFinalPageSpecificText() {
        string body = string.Join("", Enumerable.Range(1, 10).Select(page =>
            "<p" + (page == 1 ? "" : " style='break-before:page'") + ">BODY</p>"));
        HtmlRenderDocument rendered = RenderMarginSizing(120D,
            "@top-center{content:counter(page);width:max-content;border:2px solid red}"
            + "@bottom-center{content:counter(pages);width:max-content}", body: body);

        Assert.Equal(10, rendered.Pages.Count);
        Assert.Equal(10D, PageMarginTypographyText(rendered.Pages[0], "top-center").Width, 6);
        Assert.Equal(20D, PageMarginTypographyText(rendered.Pages[9], "top-center").Width, 6);
        Assert.Equal(14D, PageMarginGeometryBox(rendered.Pages[0], "top-center").Width, 6);
        Assert.Equal(24D, PageMarginGeometryBox(rendered.Pages[9], "top-center").Width, 6);
        Assert.Equal("10", PageMarginTypographyText(rendered.Pages[9], "top-center").Text);
        Assert.All(rendered.Pages, page => {
            HtmlRenderText total = PageMarginTypographyText(page, "bottom-center");
            Assert.Equal("10", total.Text);
            Assert.Equal(20D, total.Width, 6);
            HtmlRenderSemanticGroup current = PageMarginGeometryBox(page, "top-center");
            Assert.Equal(80D, current.X + current.Width / 2D, 6);
        });
        Assert.Equal(string.Concat(Enumerable.Repeat("BODY", 10)), PageMarginTypographyCompact(rendered.Text));
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginIntrinsicWidths_PercentageFitContentRetainsAnExplicitLoss() {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:'ALPHA BETA';width:fit-content(50%)}");

        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Source == "@page @top-center"
            && diagnostic.Code == HtmlRenderDiagnosticCodes.IntrinsicSizeUnsupported
            && diagnostic.Detail == "width=fit-content(50%)"
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        Assert.ThrowsAny<InvalidOperationException>(() => rendered.RequireNoLoss());
    }
}

using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("top", 0D)]
    [InlineData("bottom", 240D)]
    public void HtmlPageMarginGeometry_FixedContentWidthKeepsTheOuterBoxCentred(string side, double y) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@" + side + "-center{content:'ALPHA BETA';width:80px;border:2px solid red}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(page, side + "-center");
        HtmlRenderText text = PageMarginTypographyText(page, side + "-center");

        Assert.Equal(128D, box.X, 6);
        Assert.Equal(84D, box.Width, 6);
        Assert.Equal(y, box.Y, 6);
        Assert.Equal(60D, box.Height, 6);
        Assert.Equal(130D, text.X, 6);
        Assert.Equal(80D, text.Width, 6);
        Assert.Equal("ALPHA BETA", text.Text);
        Assert.True(MarginSizingPaintedText(page, text).Length > 1);
        Assert.Contains(box.Visuals.OfType<HtmlRenderShape>(), shape => shape.Shape.StrokeColor == OfficeColor.Red);
        Assert.Equal(HtmlRenderSemanticGroupRole.Artifact, box.Role);
        Assert.Equal("BODY", rendered.Text.Trim());
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlPageMarginGeometry_FixedEdgesKeepTheirWidthsAndLeaveTheRemainingGap(bool center) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-left{content:'L';width:80px}@top-right{content:'R';width:40px}"
            + (center ? "@top-center{content:'CENTER'}" : ""));
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderText left = PageMarginTypographyText(page, "top-left");
        HtmlRenderText right = PageMarginTypographyText(page, "top-right");

        Assert.Equal(20D, left.X, 6);
        Assert.Equal(80D, left.Width, 6);
        Assert.Equal(280D, right.X, 6);
        Assert.Equal(40D, right.Width, 6);
        if (center) {
            HtmlRenderText middle = PageMarginTypographyText(page, "top-center");
            Assert.Equal(140D, middle.Width, 6);
            Assert.Equal(170D, middle.X + middle.Width / 2D, 6);
        }
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("max-width:80px", 80D, 220D)]
    [InlineData("min-width:280px", 280D, 20D)]
    [InlineData("width:80px;min-width:132px;max-width:96px", 132D, 168D)]
    public void HtmlPageMarginGeometry_ConstraintsReallocateTheOtherEdge(
        string constraint, double leftWidth, double rightWidth) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-left{content:'AAA';" + constraint + "}@top-right{content:'R'}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);

        Assert.Equal(leftWidth, PageMarginTypographyText(page, "top-left").Width, 6);
        Assert.Equal(rightWidth, PageMarginTypographyText(page, "top-right").Width, 6);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginGeometry_CentreMaximumThenSideMinimumRecalculatesTheWholeStrip() {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-left{content:'AAA';min-width:128px}"
            + "@top-center{content:'TITLE CENTRE';max-width:112px}@top-right{content:'RR'}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderText center = PageMarginTypographyText(page, "top-center");

        Assert.Equal(128D, PageMarginTypographyText(page, "top-left").Width, 6);
        Assert.Equal(112D, center.Width, 6);
        Assert.Equal(94D, PageMarginTypographyText(page, "top-right").Width, 6);
        Assert.Equal(170D, center.X + center.Width / 2D, 6);
        Assert.Equal("TITLE CENTRE", center.Text);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("50%", 150D)]
    // The configured default paper is A4; authored @page size does not replace its unit reference.
    [InlineData("50cqw", 210D / 25.4D * 96D / 2D)]
    [InlineData("calc(50% - 10px)", 140D)]
    public void HtmlPageMarginGeometry_PercentagesAndPaperUnitsUseTheirOwnReferences(string width, double expected) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D, "@top-center{content:'A';width:" + width + "}");
        Assert.Equal(expected, PageMarginTypographyText(Assert.Single(rendered.Pages), "top-center").Width, 6);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginGeometry_AsymmetricPaddingAndBordersInsetOneContentFrame() {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:'A';width:80px;padding:6px 8px 10px 20px;"
            + "border-style:solid;border-width:2px 7px 4px 3px;border-color:red green blue black}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(page, "top-center");
        HtmlRenderText text = PageMarginTypographyText(page, "top-center");

        Assert.Equal(111D, box.X, 6);
        Assert.Equal(118D, box.Width, 6);
        Assert.Equal(60D, box.Height, 6);
        Assert.Equal(134D, text.X, 6);
        Assert.Equal(80D, text.Width, 6);
        Assert.Equal(17D, text.LayoutY, 6);
        Assert.Equal(4, box.Visuals.OfType<HtmlRenderShape>().Count());
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginGeometry_VerticalPaddingPercentagesUseTheReservedStripHeight() {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:'A';width:100px;padding:10% 2.5% 5% 5%}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(page, "top-center");
        HtmlRenderText text = PageMarginTypographyText(page, "top-center");

        Assert.Equal(122.5D, box.Width, 6);
        Assert.Equal(box.X + 15D, text.X, 6);
        Assert.Equal(100D, text.Width, 6);
        Assert.Equal(21.5D, text.LayoutY, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(0)]
    [InlineData(80)]
    public void HtmlPageMarginGeometry_EmptyGeneratedContentPaintsItsBorderWhileNoneDoesNot(int width) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:'';width:" + width + "px;border:2px solid red}"
            + "@bottom-center{content:none;width:80px;border:2px solid red}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(page, "top-center");

        Assert.Equal(width + 4D, box.Width, 6);
        Assert.Equal(60D, box.Height, 6);
        Assert.Contains(box.Visuals.OfType<HtmlRenderShape>(), shape => shape.Shape.StrokeColor == OfficeColor.Red);
        Assert.DoesNotContain(page.Scene.OfType<HtmlRenderSemanticGroup>(), group => group.Source == "@page @bottom-center");
        Assert.Empty(EnumerateRenderVisuals(box.Visuals).OfType<HtmlRenderText>());
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("0")]
    [InlineData("calc(1px - 2px)")]
    public void HtmlPageMarginGeometry_ZeroContentWidthReportsTheActualPaintLossWithoutThrowing(string width) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:'A';width:" + width + ";border:2px solid red}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderClipGroup clip = Assert.Single(EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderClipGroup>(),
            group => group.Source == "@page @top-center");

        Assert.Equal(0D, clip.ClipWidth, 6);
        Assert.True(clip.ClipHorizontal);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Source == "@page @top-center"
            && diagnostic.Detail == "margin-width" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
    }

    [Fact]
    public void HtmlPageMarginGeometry_ShorthandsVariablesAndResetUseTheExistingCascade() {
        HtmlRenderDocument styled = RenderMarginSizing(300D,
            "@top-center{content:'A';--w:80px;width:var(--w);padding:5px;padding-left:11px;"
            + "border:2px solid red;border-left-width:4px}");
        HtmlRenderPage page = Assert.Single(styled.Pages);
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(page, "top-center");
        Assert.Equal(102D, box.Width, 6);
        Assert.Equal(box.X + 15D, PageMarginTypographyText(page, "top-center").X, 6);
        styled.RequireNoLoss();

        HtmlRenderPage reset = Assert.Single(RenderMarginSizing(300D,
            "@top-center{content:'A';width:80px;width:initial}").Pages);
        HtmlRenderPage original = Assert.Single(RenderMarginSizing(300D, "@top-center{content:'A'}").Pages);
        Assert.Equal(PageMarginGeometryBox(original, "top-center").Height, PageMarginGeometryBox(reset, "top-center").Height, 6);
        Assert.Equal(PageMarginTypographyText(original, "top-center").Width, PageMarginTypographyText(reset, "top-center").Width, 6);
    }

    [Fact]
    public void HtmlPageMarginGeometry_PageSpecificWidthsAndCountersRemainArtifacts() {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:counter(page);width:80px}@bottom-center{content:counter(pages);width:60px}"
            + "}@page :left{@top-center{width:120px}",
            body: "<p>FIRST</p><p style='break-before:page'>SECOND</p>");
        Assert.Equal(2, rendered.Pages.Count);
        Assert.Equal(80D, PageMarginTypographyText(rendered.Pages[0], "top-center").Width, 6);
        Assert.Equal(120D, PageMarginTypographyText(rendered.Pages[1], "top-center").Width, 6);
        Assert.Equal("1", PageMarginTypographyText(rendered.Pages[0], "top-center").Text);
        Assert.Equal("2", PageMarginTypographyText(rendered.Pages[1], "top-center").Text);
        Assert.All(rendered.Pages, page => Assert.Equal("2", PageMarginTypographyText(page).Text));
        Assert.Equal("FIRSTSECOND", PageMarginTypographyCompact(rendered.Text));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("width:fit-content(80px)", HtmlRenderDiagnosticCodes.IntrinsicSizeUnsupported)]
    [InlineData("border:2px dashed red", HtmlRenderDiagnosticCodes.BorderPaintValueUnsupported)]
    public void HtmlPageMarginGeometry_UnsupportedDimensionsAndBorderProfilesRetainTypedLoss(
        string declaration, string code) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D, "@top-center{content:'A';" + declaration + "}");
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Source == "@page @top-center"
            && diagnostic.Code == code && diagnostic.LossKind != OfficeConversionLossKind.None);
        Assert.ThrowsAny<InvalidOperationException>(() => rendered.RequireNoLoss());
    }

    [Fact]
    public void HtmlPageMarginGeometry_AnOversizedPaintedBorderRetainsItsWidthAndReportsPageClipping() {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:'';width:400px;border:2px solid red}");
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(Assert.Single(rendered.Pages), "top-center");
        Assert.Equal(404D, box.Width, 6);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Source == "@page @top-center"
            && diagnostic.Detail == "margin-page-bounds" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.ThrowsAny<InvalidOperationException>(() => rendered.RequireNoLoss());
    }

    [Fact]
    public void HtmlPageMarginGeometry_RunningContentRelayoutUsesThePaddedInnerFrame() {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:element(heading);width:80px;padding:0 10px;border:2px solid red}",
            body: "<p style='position:running(heading);background:blue'>HEADER</p><p>BODY</p>");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup box = Assert.Single(page.Scene.OfType<HtmlRenderSemanticGroup>(),
            group => group.Source == "@page @top-center element(heading)");
        HtmlRenderText text = Assert.Single(EnumerateRenderVisuals(box.Visuals).OfType<HtmlRenderText>());

        Assert.Equal(104D, box.Width, 6);
        Assert.Equal(60D, box.Height, 6);
        Assert.Equal("HEADER", text.Text);
        Assert.Equal(box.X + 12D, text.X, 6);
        Assert.InRange(text.Width, 0.01D, 80D);
        HtmlRenderShape background = Assert.Single(EnumerateRenderVisuals(box.Visuals).OfType<HtmlRenderShape>(),
            visual => visual.Shape.FillColor == OfficeColor.Blue);
        Assert.Equal(box.X + 12D, background.X, 6);
        Assert.Equal(80D, background.Width, 6);
        Assert.Equal("BODY", rendered.Text.Trim());
        rendered.RequireNoLoss();
    }

    private static HtmlRenderSemanticGroup PageMarginGeometryBox(HtmlRenderPage page, string position) =>
        Assert.Single(page.Scene.OfType<HtmlRenderSemanticGroup>(), group => group.Source == "@page @" + position);
}

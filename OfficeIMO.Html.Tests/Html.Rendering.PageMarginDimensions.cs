using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("20px", 20D)]
    [InlineData("50%", 30D)]
    [InlineData("calc(50% - 10px)", 20D)]
    [InlineData("2ch", 20D)]
    public void HtmlPageMarginDimensions_AuthoredContentHeightKeepsTheInsideEdgeAtTheBody(
        string height, double expectedHeight) {
        HtmlRenderDocument rendered = RenderPageMarginDimensions("height:" + height, "''");

        AssertPageMarginDimensionFrames(rendered, 60D - expectedHeight, 240D, expectedHeight);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("max-height:30px", 30D)]
    [InlineData("min-height:75px", 75D)]
    [InlineData("height:20px;max-height:12px", 12D)]
    [InlineData("height:20px;min-height:32px;max-height:24px", 32D)]
    [InlineData("height:20px;min-height:50%;max-height:20%", 30D)]
    public void HtmlPageMarginDimensions_HeightConstraintsRecalculatePlacementWithMinimumPrecedence(
        string declarations, double expectedHeight) {
        HtmlRenderDocument rendered = RenderPageMarginDimensions(declarations, "''");

        AssertPageMarginDimensionFrames(rendered, 60D - expectedHeight, 240D, expectedHeight);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("5px", "15px", 5D, 40D)]
    [InlineData("10%", "5%", 6D, 51D)]
    [InlineData("calc(10% + 4px)", "1ch", 10D, 40D)]
    [InlineData("-5px", "15px", -5D, 50D)]
    public void HtmlPageMarginDimensions_AutomaticHeightSubtractsSignedVerticalMargins(
        string top, string bottom, double usedTop, double expectedHeight) {
        HtmlRenderDocument rendered = RenderPageMarginDimensions(
            "margin-top:" + top + ";margin-bottom:" + bottom);

        AssertPageMarginDimensionFrames(rendered, usedTop, 240D + usedTop, expectedHeight);
        Assert.Equal("BODY", rendered.Text.Trim());
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("5px", "15px", 25D, 245D)]
    [InlineData("auto", "15px", 25D, 265D)]
    [InlineData("5px", "auto", 5D, 245D)]
    [InlineData("auto", "auto", 20D, 260D)]
    public void HtmlPageMarginDimensions_FixedHeightSolvesVerticalAutoMarginsAndTheOutsideMargin(
        string top, string bottom, double expectedTopY, double expectedBottomY) {
        HtmlRenderDocument rendered = RenderPageMarginDimensions(
            "height:20px;margin-top:" + top + ";margin-bottom:" + bottom, "''");

        AssertPageMarginDimensionFrames(rendered, expectedTopY, expectedBottomY, 20D);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginDimensions_ContentHeightAddsAsymmetricPaddingAndBordersOnce() {
        HtmlRenderDocument rendered = RenderPageMarginDimensions(
            "width:80px;height:20px;padding:6px 8px 10px 20px;"
            + "border-style:solid;border-width:2px 7px 4px 3px;border-color:red green blue black");
        HtmlRenderPage page = Assert.Single(rendered.Pages);

        AssertPageMarginDimensionFrames(rendered, 18D, 240D, 42D);
        foreach (string position in new[] { "top-center", "bottom-center" }) {
            HtmlRenderSemanticGroup box = PageMarginGeometryBox(page, position);
            HtmlRenderText text = PageMarginTypographyText(page, position);
            Assert.Equal(118D, box.Width, 6);
            Assert.Equal(box.X + 23D, text.X, 6);
            Assert.Equal(80D, text.Width, 6);
            Assert.Equal(box.Y + 8D, text.LayoutY, 6);
            Assert.Equal(4, box.Visuals.OfType<HtmlRenderShape>().Count());
        }
        Assert.Equal(2, PageMarginTypographyDrawingText(page.CreateDrawing()).Count(text => text.Text == "A"));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("10px", "30px", 30D, 180D)]
    [InlineData("5%", "10%", 35D, 175D)]
    [InlineData("-10px", "20px", 10D, 210D)]
    [InlineData("-100px", "-100px", -80D, 420D)]
    public void HtmlPageMarginDimensions_SignedHorizontalMarginFramesAllocateTheRemainingPeerWidth(
        string leftMargin, string rightMargin, double expectedLeftX, double expectedRightWidth) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-left{content:'';width:80px;height:20px;margin-left:" + leftMargin
            + ";margin-right:" + rightMargin + "}@top-right{content:'R';height:20px}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup left = PageMarginGeometryBox(page, "top-left");
        HtmlRenderText right = PageMarginTypographyText(page, "top-right");

        Assert.Equal(expectedLeftX, left.X, 6);
        Assert.Equal(80D, left.Width, 6);
        Assert.Equal(expectedRightWidth, right.Width, 6);
        Assert.Equal(320D, right.X + right.Width, 6);
        Assert.Equal("R", right.Text);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("10px", "30px", 120D)]
    [InlineData("-10px", "20px", 115D)]
    public void HtmlPageMarginDimensions_CentreAlignmentUsesTheAsymmetricMarginFrame(
        string leftMargin, string rightMargin, double expectedX) {
        HtmlRenderDocument rendered = RenderPageMarginDimensions(
            "width:80px;height:20px;margin-left:" + leftMargin + ";margin-right:" + rightMargin, "''");
        HtmlRenderPage page = Assert.Single(rendered.Pages);

        foreach (string position in new[] { "top-center", "bottom-center" }) {
            HtmlRenderSemanticGroup box = PageMarginGeometryBox(page, position);
            Assert.Equal(expectedX, box.X, 6);
            Assert.Equal(80D, box.Width, 6);
        }
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("auto", "auto", 20D, 220D)]
    [InlineData("auto", "30px", 20D, 190D)]
    [InlineData("10px", "auto", 30D, 210D)]
    public void HtmlPageMarginDimensions_HorizontalAutoMarginsUseZeroBeforePeerAllocation(
        string leftMargin, string rightMargin, double expectedX, double expectedPeerWidth) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-left{content:'';width:80px;height:20px;margin-left:" + leftMargin
            + ";margin-right:" + rightMargin + "}@top-right{content:'R';height:20px}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);

        Assert.Equal(expectedX, PageMarginGeometryBox(page, "top-left").X, 6);
        Assert.Equal(expectedPeerWidth, PageMarginTypographyText(page, "top-right").Width, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("none")]
    [InlineData("normal")]
    public void HtmlPageMarginDimensions_EmptyContentKeepsAuthoredFramesWhileAbsentContentDoesNot(string absent) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:'';height:20px;border:2px solid red}"
            + "@top-left{content:'';margin:5px}"
            + "@bottom-center{content:" + absent + ";height:20px;margin:5px;border:2px solid red}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup decorated = PageMarginGeometryBox(page, "top-center");
        HtmlRenderSemanticGroup marginsOnly = PageMarginGeometryBox(page, "top-left");

        Assert.Equal(36D, decorated.Y, 6);
        Assert.Equal(24D, decorated.Height, 6);
        Assert.Contains(decorated.Visuals.OfType<HtmlRenderShape>(), shape => shape.Shape.StrokeColor == OfficeColor.Red);
        Assert.Equal(5D, marginsOnly.Y, 6);
        Assert.Equal(50D, marginsOnly.Height, 6);
        Assert.Empty(EnumerateRenderVisuals(decorated.Visuals).OfType<HtmlRenderText>());
        Assert.DoesNotContain(page.Scene.OfType<HtmlRenderSemanticGroup>(), group => group.Source == "@page @bottom-center");
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("height:0")]
    [InlineData("height:calc(1px - 2px)")]
    [InlineData("margin-top:40px;margin-bottom:40px")]
    public void HtmlPageMarginDimensions_ZeroContentHeightRetainsTextBehindAnExactClipAndTypedLoss(string declarations) {
        HtmlRenderDocument rendered = RenderPageMarginDimensions(declarations);
        HtmlRenderPage page = Assert.Single(rendered.Pages);

        foreach (string position in new[] { "top-center", "bottom-center" }) {
            string source = "@page @" + position;
            HtmlRenderClipGroup clip = Assert.Single(EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderClipGroup>(),
                candidate => candidate.Source == source);
            Assert.Equal(0D, clip.ClipHeight, 6);
            Assert.True(clip.ClipVertical);
            Assert.Equal("A", PageMarginTypographyText(page, position).Text);
            Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Source == source
                && diagnostic.Code == HtmlRenderDiagnosticCodes.GeneratedContentUnsupported
                && diagnostic.Detail == "margin-height" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        }
        Assert.ThrowsAny<InvalidOperationException>(() => rendered.RequireNoLoss());
    }

    [Theory]
    [InlineData("", 80D, false)]
    [InlineData("border:2px solid red", 84D, true)]
    public void HtmlPageMarginDimensions_OversizedEmptyFramesOnlyReportPhysicalLossForPaintedBorders(
        string decoration, double expectedHeight, bool expectedLoss) {
        HtmlRenderDocument rendered = RenderPageMarginDimensions("height:80px;" + decoration, "''");
        HtmlRenderPage page = Assert.Single(rendered.Pages);

        AssertPageMarginDimensionFrames(rendered, 60D - expectedHeight, 240D, expectedHeight);
        Assert.Equal(expectedLoss, rendered.HasLoss);
        if (expectedLoss) {
            foreach (string position in new[] { "top-center", "bottom-center" }) {
                Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Source == "@page @" + position
                    && diagnostic.Detail == "margin-page-bounds" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
            }
            Assert.ThrowsAny<InvalidOperationException>(() => rendered.RequireNoLoss());
        } else rendered.RequireNoLoss();
        Assert.DoesNotContain(EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderText>(), text => text.SemanticRole == "page-margin");
    }

    [Theory]
    [InlineData("height:min-content")]
    [InlineData("min-height:max-content")]
    [InlineData("max-height:fit-content(20px)")]
    public void HtmlPageMarginDimensions_IntrinsicHeightFallbackReportsTheUnsupportedConstraint(string declaration) {
        HtmlRenderDocument rendered = RenderPageMarginDimensions("width:80px;" + declaration, "''");

        AssertPageMarginDimensionFrames(rendered, 0D, 240D, 60D);
        foreach (string position in new[] { "top-center", "bottom-center" }) {
            Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Source == "@page @" + position
                && diagnostic.Code == HtmlRenderDiagnosticCodes.IntrinsicSizeUnsupported
                && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        }
        Assert.ThrowsAny<InvalidOperationException>(() => rendered.RequireNoLoss());
    }

    [Fact]
    public void HtmlPageMarginDimensions_VariablesAndMarginShorthandKeepLonghandOrder() {
        HtmlRenderDocument rendered = RenderPageMarginDimensions(
            "--h:20px;--gap:5px;height:var(--h);margin:var(--gap);margin-bottom:15px", "''");
        HtmlRenderPage page = Assert.Single(rendered.Pages);

        AssertPageMarginDimensionFrames(rendered, 25D, 245D, 20D);
        Assert.Equal(25D, PageMarginGeometryBox(page, "top-center").X, 6);
        Assert.Equal(290D, PageMarginGeometryBox(page, "top-center").Width, 6);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginDimensions_RunningContentUsesTheConstrainedPaddedInnerFrame() {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:element(heading);width:80px;height:32px;padding:6px 0;border:2px solid red}",
            body: "<p style='position:running(heading);height:20px;background:blue'>HEADER</p><p>BODY</p>");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup box = Assert.Single(page.Scene.OfType<HtmlRenderSemanticGroup>(),
            group => group.Source == "@page @top-center element(heading)");
        HtmlRenderShape background = Assert.Single(EnumerateRenderVisuals(box.Visuals).OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillColor == OfficeColor.Blue);

        Assert.Equal(84D, box.Width, 6);
        Assert.Equal(48D, box.Height, 6);
        Assert.Equal(12D, box.Y, 6);
        Assert.Equal(box.X + 2D, background.X, 6);
        Assert.Equal(26D, background.Y, 6);
        Assert.Equal(80D, background.Width, 6);
        Assert.Equal(20D, background.Height, 6);
        Assert.Equal("HEADER", Assert.Single(EnumerateRenderVisuals(box.Visuals).OfType<HtmlRenderText>()).Text);
        Assert.Contains(PageMarginTypographyDrawingText(page.CreateDrawing()), text => text.Text == "HEADER");
        Assert.Equal("BODY", rendered.Text.Trim());
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginDimensions_ZeroHeightRunningContentKeepsTheSharedExactClipAndLoss() {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:element(heading);width:80px;height:0;padding:6px 0;border:2px solid red}",
            body: "<p style='position:running(heading)'>HEADER</p><p>BODY</p>");
        HtmlRenderSemanticGroup box = Assert.Single(Assert.Single(rendered.Pages).Scene.OfType<HtmlRenderSemanticGroup>(),
            group => group.Source == "@page @top-center element(heading)");
        HtmlRenderClipGroup clip = Assert.Single(box.Visuals.OfType<HtmlRenderClipGroup>());

        Assert.Equal(16D, box.Height, 6);
        Assert.Equal(44D, box.Y, 6);
        Assert.Equal(0D, clip.ClipHeight, 6);
        Assert.True(clip.ClipVertical);
        Assert.Equal("HEADER", Assert.Single(EnumerateRenderVisuals(clip.Visuals).OfType<HtmlRenderText>()).Text);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Source == box.Source
            && diagnostic.Detail == "margin-height" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.Equal("BODY", rendered.Text.Trim());
        Assert.ThrowsAny<InvalidOperationException>(() => rendered.RequireNoLoss());
    }

    [Theory]
    [InlineData("width:80px;height:20px", "height:40px", "margin-height")]
    [InlineData("width:40px;height:40px", "width:80px;height:20px", "margin-width")]
    public void HtmlPageMarginDimensions_PositiveRunningClipReportsPaintOmission(
        string boxStyle, string runningStyle, string expectedDetail) {
        string running = expectedDetail == "margin-width"
            ? "<table style='position:running(heading);display:inline-table;border-collapse:collapse;background:blue;"
                + runningStyle + "'><tr><td style='padding:0'>HEADER</td></tr></table>"
            : "<p style='position:running(heading);background:blue;" + runningStyle + "'>HEADER</p>";
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:element(heading);" + boxStyle + "}"
            + "@bottom-center{content:element(heading);" + boxStyle + "}",
            body: running + "<p>BODY</p>");
        HtmlRenderPage page = Assert.Single(rendered.Pages);

        foreach (string position in new[] { "top-center", "bottom-center" }) {
            string source = "@page @" + position + " element(heading)";
            HtmlRenderSemanticGroup box = Assert.Single(page.Scene.OfType<HtmlRenderSemanticGroup>(), group => group.Source == source);
            HtmlRenderClipGroup clip = Assert.Single(box.Visuals.OfType<HtmlRenderClipGroup>());
            Assert.True(clip.ClipWidth > 0D && clip.ClipHeight > 0D);
            Assert.Equal(expectedDetail == "margin-width", clip.ClipHorizontal);
            Assert.Equal(expectedDetail == "margin-height", clip.ClipVertical);
            Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Source == source
                && diagnostic.Detail == expectedDetail && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        }
        Assert.Equal("BODY", rendered.Text.Trim());
        Assert.ThrowsAny<InvalidOperationException>(() => rendered.RequireNoLoss());
    }

    [Fact]
    public void HtmlPageMarginDimensions_EmptyRunningFrameOverflowDoesNotInventPaintLoss() {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:element(heading);width:80px;height:40px;text-align:left}",
            body: "<table style='position:running(heading);display:inline-table;border-collapse:collapse;width:160px;height:20px'>"
                + "<tr><td style='padding:0'>OK</td></tr></table><p>BODY</p>");
        HtmlRenderSemanticGroup box = Assert.Single(Assert.Single(rendered.Pages).Scene.OfType<HtmlRenderSemanticGroup>(),
            group => group.Source == "@page @top-center element(heading)");
        HtmlRenderClipGroup clip = Assert.Single(box.Visuals.OfType<HtmlRenderClipGroup>());

        Assert.True(clip.ClipHorizontal);
        Assert.False(clip.ClipVertical);
        Assert.Equal("OK", Assert.Single(EnumerateRenderVisuals(clip.Visuals).OfType<HtmlRenderText>()).Text);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginDimensions_PageSpecificPercentagesUseThatPagesGeometryAndFinalCounters() {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:counter(page) '/' counter(pages);width:80px;height:50%;"
            + "margin-top:10%;margin-bottom:auto}}@page :first{margin-top:40px}@page :left{margin-top:80px",
            body: "<p>FIRST</p><p style='break-before:page'>SECOND</p>");

        Assert.Equal(2, rendered.Pages.Count);
        HtmlRenderSemanticGroup first = PageMarginGeometryBox(rendered.Pages[0], "top-center");
        HtmlRenderSemanticGroup second = PageMarginGeometryBox(rendered.Pages[1], "top-center");
        Assert.Equal(20D, first.Height, 6);
        Assert.Equal(4D, first.Y, 6);
        Assert.Equal(40D, second.Height, 6);
        Assert.Equal(8D, second.Y, 6);
        Assert.Equal("1/2", PageMarginTypographyText(rendered.Pages[0], "top-center").Text);
        Assert.Equal("2/2", PageMarginTypographyText(rendered.Pages[1], "top-center").Text);
        Assert.Equal("FIRSTSECOND", PageMarginTypographyCompact(rendered.Text));
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginDimensions_FinalCountersRetainCompleteSceneTextInsideTheAuthoredHeightClip() {
        string body = string.Join("", Enumerable.Range(1, 10).Select(page =>
            "<p" + (page == 1 ? "" : " style='break-before:page'") + ">BODY</p>"));
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:counter(page) '\\A' counter(pages);width:80px;height:20px}", body: body);

        Assert.Equal(10, rendered.Pages.Count);
        Assert.Equal("110", PageMarginTypographyCompact(PageMarginTypographyText(rendered.Pages[0], "top-center").Text));
        Assert.Equal("1010", PageMarginTypographyCompact(PageMarginTypographyText(rendered.Pages[9], "top-center").Text));
        Assert.All(rendered.Pages, page => {
            HtmlRenderClipGroup clip = Assert.Single(EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderClipGroup>(),
                group => group.Source == "@page @top-center");
            Assert.Equal(20D, clip.ClipHeight, 6);
            Assert.True(clip.ClipVertical);
            Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Source == clip.Source
                && diagnostic.Detail == "margin-height" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        });
        Assert.Equal(10, rendered.Text.Split(new[] { "BODY" }, StringSplitOptions.None).Length - 1);
        Assert.ThrowsAny<InvalidOperationException>(() => rendered.RequireNoLoss());
    }

    private static HtmlRenderDocument RenderPageMarginDimensions(string declarations, string content = "'A'") =>
        RenderMarginSizing(300D, "@top-center{content:" + content + ";" + declarations
            + "}@bottom-center{content:" + content + ";" + declarations + "}");

    private static void AssertPageMarginDimensionFrames(HtmlRenderDocument rendered,
        double expectedTopY, double expectedBottomY, double expectedHeight) {
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup top = PageMarginGeometryBox(page, "top-center");
        HtmlRenderSemanticGroup bottom = PageMarginGeometryBox(page, "bottom-center");
        Assert.Equal(expectedTopY, top.Y, 6);
        Assert.Equal(expectedBottomY, bottom.Y, 6);
        Assert.Equal(expectedHeight, top.Height, 6);
        Assert.Equal(expectedHeight, bottom.Height, 6);
        Assert.Equal(HtmlRenderSemanticGroupRole.Artifact, top.Role);
        Assert.Equal(HtmlRenderSemanticGroupRole.Artifact, bottom.Role);
    }
}

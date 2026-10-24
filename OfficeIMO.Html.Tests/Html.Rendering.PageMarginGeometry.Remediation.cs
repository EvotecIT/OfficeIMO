using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlPageMarginGeometryRemediation_PaddingExhaustingTheStripDiagnosesClippedRunningText() {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:element(heading);width:80px;padding:30px 0}",
            body: "<p style='position:running(heading)'>HEADER</p><p>BODY</p>");
        HtmlRenderSemanticGroup box = Assert.Single(Assert.Single(rendered.Pages).Scene.OfType<HtmlRenderSemanticGroup>(),
            group => group.Source == "@page @top-center element(heading)");

        Assert.Equal(60D, box.Height, 6);
        Assert.Equal(0D, Assert.Single(box.Visuals.OfType<HtmlRenderClipGroup>()).ClipHeight);
        Assert.Equal("HEADER", Assert.Single(EnumerateRenderVisuals(box.Visuals).OfType<HtmlRenderText>()).Text);
        Assert.Equal("BODY", rendered.Text.Trim());
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Source == box.Source
            && diagnostic.Detail == "margin-height" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Detail == "margin-page-bounds");
        Assert.ThrowsAny<InvalidOperationException>(() => rendered.RequireNoLoss());
    }

    [Theory]
    [InlineData("'PAD'", "", true)]
    [InlineData("''", "padding:10px 0;border:2px solid red", true)]
    [InlineData("''", "", false)]
    public void HtmlPageMarginGeometryRemediation_ZeroReservedStripRetainsAuthoredGeometryAndMeaningfulLoss(
        string content, string decoration, bool expectedLoss) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "margin-top:0;@top-center{content:" + content + ";width:80px;" + decoration + "}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(page, "top-center");

        Assert.Equal(0D, page.Margins.Top);
        Assert.Equal(expectedLoss, rendered.HasLoss);
        if (decoration.Length > 0) {
            Assert.Equal(24D, box.Height, 6);
            Assert.Equal(-24D, box.Y, 6);
            Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Detail == "margin-page-bounds"
                && diagnostic.Source == box.Source && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        } else if (content == "'PAD'") {
            Assert.Equal(0D, Assert.Single(box.Visuals.OfType<HtmlRenderClipGroup>()).ClipHeight);
            Assert.Equal("PAD", Assert.Single(EnumerateRenderVisuals(box.Visuals).OfType<HtmlRenderText>()).Text);
            Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Detail == "margin-height"
                && diagnostic.Source == box.Source && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        } else Assert.Empty(box.Visuals);
        if (expectedLoss) Assert.ThrowsAny<InvalidOperationException>(() => rendered.RequireNoLoss());
        else rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("top", 70, 2, 144D, -84D)]
    [InlineData("bottom", 50, 2, 104D, 240D)]
    [InlineData("top", 0, 80, 160D, -100D)]
    public void HtmlPageMarginGeometryRemediation_InsetsExpandFromTheBodyEdgeAndReportPhysicalBorderClipping(
        string side, int padding, int border, double expectedHeight, double expectedY) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@" + side + "-center{content:'PAD';width:80px;padding:" + padding
            + "px 0;border:" + border + "px solid red}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(page, side + "-center");

        Assert.Equal(expectedHeight, box.Height, 6);
        Assert.Equal(expectedY, box.Y, 6);
        Assert.Equal(80D + border * 2D, box.Width, 6);
        if (side == "top") Assert.Equal(page.Margins.Top, box.Y + box.Height, 6);
        else Assert.Equal(page.Height - page.Margins.Bottom, box.Y, 6);
        Assert.Contains(EnumerateRenderVisuals(box.Visuals).OfType<HtmlRenderShape>(),
            visual => visual.Shape.StrokeColor == OfficeColor.Red);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Source == "@page @" + side + "-center"
            && diagnostic.Detail == "margin-page-bounds" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.ThrowsAny<InvalidOperationException>(() => rendered.RequireNoLoss());
    }

    [Theory]
    [InlineData("left", true)]
    [InlineData("center", false)]
    [InlineData("right", true)]
    public void HtmlPageMarginGeometryRemediation_BorderlessWideTextDiagnosesActualPaintInsteadOfItsFrame(
        string alignment, bool expectedLoss) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:'HEADER';width:1000px;text-align:" + alignment + "}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);

        Assert.Equal(1000D, PageMarginGeometryBox(page, "top-center").Width, 6);
        Assert.Equal("HEADER", PageMarginTypographyText(page, "top-center").Text);
        Assert.Equal(expectedLoss, rendered.HasLoss);
        if (expectedLoss) {
            Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Source == "@page @top-center"
                && diagnostic.Detail == "margin-page-bounds" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
            Assert.ThrowsAny<InvalidOperationException>(() => rendered.RequireNoLoss());
        } else rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("left", true)]
    [InlineData("center", false)]
    [InlineData("right", true)]
    public void HtmlPageMarginGeometryRemediation_RunningTextChecksTranslatedPhysicalPaint(
        string alignment, bool expectedLoss) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:element(heading);width:1000px}",
            body: "<p style='position:running(heading);text-align:" + alignment + "'>HEADER</p><p>BODY</p>");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup box = Assert.Single(page.Scene.OfType<HtmlRenderSemanticGroup>(),
            group => group.Source == "@page @top-center element(heading)");

        Assert.Equal(1000D, box.Width, 6);
        Assert.Equal("HEADER", Assert.Single(EnumerateRenderVisuals(box.Visuals).OfType<HtmlRenderText>()).Text);
        Assert.Equal("BODY", rendered.Text.Trim());
        Assert.Equal(expectedLoss, rendered.HasLoss);
        if (expectedLoss) {
            Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Source == box.Source
                && diagnostic.Detail == "margin-page-bounds" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
            Assert.ThrowsAny<InvalidOperationException>(() => rendered.RequireNoLoss());
        } else rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("color:transparent")]
    [InlineData("opacity:0")]
    public void HtmlPageMarginGeometryRemediation_InvisibleRunningTextDoesNotReportPhysicalClipping(string paintStyle) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:element(heading);width:1000px}",
            body: "<p style='position:running(heading);text-align:left;" + paintStyle + "'>HEADER</p><p>BODY</p>");

        Assert.Equal("BODY", rendered.Text.Trim());
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("none", false)]
    [InlineData("normal", false)]
    [InlineData("''", true)]
    public void HtmlPageMarginGeometryRemediation_OnlyGeneratedContentHasApplicableGeometryFallbacks(
        string content, bool expectedLoss) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:" + content + ";width:fit-content(80px)}");

        Assert.Equal(expectedLoss, rendered.HasLoss);
        if (expectedLoss) {
            Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Source == "@page @top-center"
                && diagnostic.Code == HtmlRenderDiagnosticCodes.IntrinsicSizeUnsupported);
            Assert.ThrowsAny<InvalidOperationException>(() => rendered.RequireNoLoss());
        } else rendered.RequireNoLoss();
    }
}

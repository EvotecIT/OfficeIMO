using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("min-content")]
    [InlineData("max-content")]
    [InlineData("fit-content")]
    public void HtmlPageMarginIntrinsicHeights_KeywordsUseTheContentWrappedAtTheAssignedWidth(string keyword) {
        HtmlRenderDocument rendered = RenderPageMarginDimensions("width:50px;height:" + keyword, "'ALPHA BETA'");
        HtmlRenderPage page = Assert.Single(rendered.Pages);

        AssertPageMarginDimensionFrames(rendered, 20D, 240D, 40D);
        foreach (string position in new[] { "top-center", "bottom-center" }) {
            HtmlRenderText text = PageMarginTypographyText(page, position);
            Assert.Equal(50D, text.Width, 6);
            Assert.Equal(40D, text.LayoutHeight, 6);
            Assert.Equal(new[] { "ALPHA", "BETA" }, IntrinsicHeightPaintedText(page, text).Select(line => line.Text).ToArray());
            Assert.Equal("ALPHA BETA", text.Text);
        }
        Assert.Equal("BODY", rendered.Text.Trim());
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("height:10px;min-height:max-content", 20D, 40D, 240D)]
    [InlineData("height:50px;max-height:min-content", 20D, 40D, 240D)]
    [InlineData("height:5px;min-height:max-content;max-height:10px", 20D, 40D, 240D)]
    [InlineData("min-height:max-content;margin-top:5px;margin-bottom:10px", 45D, 5D, 245D)]
    [InlineData("max-height:fit-content;margin-top:5px;margin-bottom:10px", 20D, 30D, 245D)]
    [InlineData("height:min-content;min-height:30px;max-height:10px", 30D, 30D, 240D)]
    public void HtmlPageMarginIntrinsicHeights_ConstraintsPreserveMinimumPrecedenceAndTheAutomaticStripResidual(
        string declarations, double expectedHeight, double expectedTopY, double expectedBottomY) {
        HtmlRenderDocument rendered = RenderPageMarginDimensions(declarations);

        AssertPageMarginDimensionFrames(rendered, expectedTopY, expectedBottomY, expectedHeight);
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        foreach (string position in new[] { "top-center", "bottom-center" }) {
            Assert.Equal("A", PageMarginTypographyText(page, position).Text);
        }
        Assert.Equal("BODY", rendered.Text.Trim());
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("content-box")]
    [InlineData("border-box")]
    public void HtmlPageMarginIntrinsicHeights_KeywordMeasurementsKeepTheirContentEdgesAndAddInsetsOnce(string boxSizing) {
        HtmlRenderDocument rendered = RenderPageMarginDimensions(
            "box-sizing:" + boxSizing + ";width:min-content;height:max-content;padding:3px 10px 7px;border:2px solid red",
            "'ALPHA BETA'");
        HtmlRenderPage page = Assert.Single(rendered.Pages);

        AssertPageMarginDimensionFrames(rendered, 6D, 240D, 54D);
        foreach (string position in new[] { "top-center", "bottom-center" }) {
            HtmlRenderSemanticGroup box = PageMarginGeometryBox(page, position);
            HtmlRenderText text = PageMarginTypographyText(page, position);
            Assert.Equal(74D, box.Width, 6);
            Assert.Equal(50D, text.Width, 6);
            Assert.Equal(box.X + 12D, text.X, 6);
            Assert.Equal(box.Y + 5D, text.LayoutY, 6);
            Assert.Equal(40D, text.LayoutHeight, 6);
            Assert.Equal(2, IntrinsicHeightPaintedText(page, text).Length);
            Assert.Contains(box.Visuals.OfType<HtmlRenderShape>(), shape => shape.Shape.StrokeColor == OfficeColor.Red);
        }
        Assert.Equal("BODY", rendered.Text.Trim());
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginIntrinsicHeights_AnAuthoredHardLineMakesAHeightOnlyBoxKeepItsBodyFacingEdge() {
        HtmlRenderDocument rendered = RenderPageMarginDimensions("height:max-content", "'A\\A B'");
        HtmlRenderPage page = Assert.Single(rendered.Pages);

        AssertPageMarginDimensionFrames(rendered, 20D, 240D, 40D);
        foreach (string position in new[] { "top-center", "bottom-center" }) {
            HtmlRenderText text = PageMarginTypographyText(page, position);
            Assert.Equal(new[] { "A", "B" }, IntrinsicHeightPaintedText(page, text).Select(line => line.Text).ToArray());
            Assert.Equal(PageMarginGeometryBox(page, position).Y, text.LayoutY, 6);
        }
        Assert.Equal("BODY", rendered.Text.Trim());
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginIntrinsicHeights_AnEmptyStringRetainsOnlyItsInsetsAndBorderWhileNoneRemovesTheBox() {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:'';height:max-content;padding:3px;border:2px solid red}"
            + "@bottom-center{content:none;height:max-content;padding:3px;border:2px solid red}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup box = PageMarginGeometryBox(page, "top-center");

        Assert.Equal(10D, box.Height, 6);
        Assert.Equal(50D, box.Y, 6);
        Assert.Contains(box.Visuals.OfType<HtmlRenderShape>(), shape => shape.Shape.StrokeColor == OfficeColor.Red);
        Assert.Empty(EnumerateRenderVisuals(box.Visuals).OfType<HtmlRenderText>());
        Assert.DoesNotContain(page.Scene.OfType<HtmlRenderSemanticGroup>(), group => group.Source == "@page @bottom-center");
        Assert.Equal("BODY", rendered.Text.Trim());
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginIntrinsicHeights_RunningContentUsesItsHeightAfterRelayoutAtTheFinalInnerWidth() {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:element(heading);width:50px;height:max-content;padding:3px 10px 7px;border:2px solid red}",
            body: "<p style='position:running(heading)'>ALPHA BETA</p><p>BODY</p>");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup box = Assert.Single(page.Scene.OfType<HtmlRenderSemanticGroup>(),
            group => group.Source == "@page @top-center element(heading)");
        HtmlRenderText[] lines = EnumerateRenderVisuals(box.Visuals).OfType<HtmlRenderText>().ToArray();

        Assert.Equal(74D, box.Width, 6);
        Assert.Equal(54D, box.Height, 6);
        Assert.Equal(6D, box.Y, 6);
        Assert.Equal(2, lines.Length);
        Assert.All(lines, line => {
            Assert.Equal(box.X + 12D, line.X, 6);
            Assert.InRange(line.Width, 0.01D, 50D);
        });
        Assert.Equal(box.Y + 5D, lines[0].LayoutY, 6);
        Assert.Equal(lines[0].LayoutY + 20D, lines[1].LayoutY, 6);
        Assert.Equal("ALPHABETA", PageMarginTypographyCompact(string.Concat(lines.Select(line => line.Text))));
        Assert.Equal("BODY", rendered.Text.Trim());
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginIntrinsicHeights_APaintNeutralRunningBlockRetainsItsAuthoredLayoutHeight() {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:element(heading);height:max-content;border:2px solid red}",
            body: "<p style='position:running(heading);height:24px'></p><p>BODY</p>");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup box = Assert.Single(page.Scene.OfType<HtmlRenderSemanticGroup>(),
            group => group.Source?.StartsWith("@page @top-center", StringComparison.Ordinal) == true);

        Assert.Equal(28D, box.Height, 6);
        Assert.Equal(32D, box.Y, 6);
        Assert.Contains(box.Visuals.OfType<HtmlRenderShape>(), shape => shape.Shape.StrokeColor == OfficeColor.Red);
        Assert.Empty(EnumerateRenderVisuals(box.Visuals).OfType<HtmlRenderText>());
        Assert.Equal("BODY", rendered.Text.Trim());
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginIntrinsicHeights_ANumericMaximumClipsPaintWithoutDiscardingTheMeasuredText() {
        HtmlRenderDocument rendered = RenderPageMarginDimensions(
            "width:50px;height:max-content;max-height:20px", "'ALPHA BETA'");
        HtmlRenderPage page = Assert.Single(rendered.Pages);

        AssertPageMarginDimensionFrames(rendered, 40D, 240D, 20D);
        foreach (string position in new[] { "top-center", "bottom-center" }) {
            HtmlRenderSemanticGroup box = PageMarginGeometryBox(page, position);
            HtmlRenderClipGroup clip = Assert.Single(EnumerateRenderVisuals(box.Visuals).OfType<HtmlRenderClipGroup>());
            Assert.Equal(20D, clip.ClipHeight, 6);
            Assert.True(clip.ClipVertical);
            Assert.Equal("ALPHA BETA", PageMarginTypographyText(page, position).Text);
            Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Source == box.Source
                && diagnostic.Code == HtmlRenderDiagnosticCodes.GeneratedContentUnsupported
                && diagnostic.Detail == "margin-height" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        }
        Assert.Equal("BODY", rendered.Text.Trim());
        Assert.ThrowsAny<InvalidOperationException>(() => rendered.RequireNoLoss());
    }

    [Fact]
    public void HtmlPageMarginIntrinsicHeights_FinalPageCountersCanChangeTheNaturalWrappedHeight() {
        string body = string.Join("", Enumerable.Range(1, 10).Select(page =>
            "<p" + (page == 1 ? "" : " style='break-before:page'") + ">BODY</p>"));
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-center{content:counter(page) ' A';width:30px;height:max-content}", body: body);

        Assert.Equal(10, rendered.Pages.Count);
        HtmlRenderSemanticGroup first = PageMarginGeometryBox(rendered.Pages[0], "top-center");
        HtmlRenderSemanticGroup last = PageMarginGeometryBox(rendered.Pages[9], "top-center");
        Assert.Equal(20D, first.Height, 6);
        Assert.Equal(40D, first.Y, 6);
        Assert.Equal(40D, last.Height, 6);
        Assert.Equal(20D, last.Y, 6);
        Assert.Equal("1 A", PageMarginTypographyText(rendered.Pages[0], "top-center").Text);
        Assert.Equal("10 A", PageMarginTypographyText(rendered.Pages[9], "top-center").Text);
        Assert.Single(IntrinsicHeightPaintedText(rendered.Pages[0], PageMarginTypographyText(rendered.Pages[0], "top-center")));
        Assert.Equal(2, IntrinsicHeightPaintedText(rendered.Pages[9], PageMarginTypographyText(rendered.Pages[9], "top-center")).Length);
        Assert.Equal(string.Concat(Enumerable.Repeat("BODY", 10)), PageMarginTypographyCompact(rendered.Text));
        rendered.RequireNoLoss();
    }
    [Theory]
    [InlineData("'A'", 20D)]
    [InlineData("'A\\A B'", 40D)]
    public void HtmlPageMarginIntrinsicHeights_ZeroSizeTextKeepsAuthoredLeadingWithoutPaintingGlyphs(string content, double height) {
        HtmlRenderDocument rendered = RenderPageMarginDimensions("font-size:0;line-height:20px;height:max-content;border:2px solid red", content);
        HtmlRenderPage page = Assert.Single(rendered.Pages);

        AssertPageMarginDimensionFrames(rendered, 60D - height - 4D, 240D, height + 4D);
        foreach (string position in new[] { "top-center", "bottom-center" }) {
            HtmlRenderSemanticGroup box = PageMarginGeometryBox(page, position);
            Assert.Empty(EnumerateRenderVisuals(box.Visuals).OfType<HtmlRenderText>());
            Assert.Contains(box.Visuals.OfType<HtmlRenderShape>(), shape => shape.Shape.StrokeColor == OfficeColor.Red);
        }
        Assert.Equal("BODY", rendered.Text.Trim());
        rendered.RequireNoLoss();
    }

    // Each fixture has one generated box per opposite strip and a distinct
    // BODY marker. Glyph origins may cross the exact strip boundary.
    private static OfficeDrawingText[] IntrinsicHeightPaintedText(HtmlRenderPage page, HtmlRenderText margin) =>
        PageMarginTypographyDrawingText(page.CreateDrawing()).Where(text => text.Text != "BODY"
            && (margin.Source!.StartsWith("@page @top-", StringComparison.Ordinal)
                ? text.Y < page.Height / 2D : text.Y >= page.Height / 2D)).ToArray();
}

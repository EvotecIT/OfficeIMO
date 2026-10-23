using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.TestAssets;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("left-top", 70D, 60D)]
    [InlineData("left-middle", 70D, 190D)]
    [InlineData("left-bottom", 70D, 320D)]
    [InlineData("right-top", 320D, 60D)]
    [InlineData("right-middle", 320D, 190D)]
    [InlineData("right-bottom", 320D, 320D)]
    [InlineData("top-left-corner", 70D, 20D)]
    [InlineData("top-right-corner", 320D, 20D)]
    [InlineData("bottom-left-corner", 70D, 360D)]
    [InlineData("bottom-right-corner", 320D, 360D)]
    public void HtmlPageMarginSideCorner_NumericBoxModelPaintsAllTenPositions(
        string position, double x, double y) {
        HtmlRenderDocument rendered = RenderSideCornerMargins("@" + position
            + "{content:'A';width:20px;height:30px;padding:3px;border:2px solid green;background:blue}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderShape fill = SideCornerMarginFill(page, OfficeColor.Blue);
        AssertSideCornerMarginRectangle(fill, x, y, 30D, 40D);
        HtmlRenderText text = PageMarginTypographyText(page, position);
        Assert.Equal(x + 5D, text.X, 6);
        Assert.Equal(20D, text.Width, 6);
        Assert.Equal("A", text.Text);
        OfficeRasterImage pixels = OfficeDrawingRasterRenderer.Render(page.CreateDrawing());
        Assert.Equal(OfficeColor.Green, pixels.GetPixel((int)x + 1, (int)y + 1));
        Assert.Equal(OfficeColor.Blue, pixels.GetPixel((int)x + 3, (int)y + 3));
        Assert.Equal(OfficeColor.White, pixels.GetPixel((int)x - 1, (int)y + 3));
        Assert.Equal("BODY", rendered.Text.Trim());
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("left-middle", "margin-left:3px;margin-right:7px", 40, 53D, false)]
    [InlineData("right-middle", "margin-left:7px;margin-right:3px", 40, 327D, false)]
    [InlineData("left-middle", "margin-left:3px;margin-right:7px", 120, -27D, true)]
    [InlineData("right-middle", "margin-left:7px;margin-right:3px", 120, 327D, true)]
    public void HtmlPageMarginSideCorner_FixedSideWidthPreservesTheBodyFacingMargin(
        string position, string margins, int width, double x, bool clipped) {
        HtmlRenderDocument rendered = RenderSideCornerMargins("@" + position
            + "{content:'';width:" + width + "px;height:20px;background:blue;" + margins + "}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        AssertSideCornerMarginRectangle(SideCornerMarginFill(page, OfficeColor.Blue), x, 200D, width, 20D);
        if (clipped) {
            Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Detail == "margin-page-bounds"
                && diagnostic.Source == "@page @" + position && diagnostic.LossKind == OfficeConversionLossKind.Omission);
            Assert.ThrowsAny<InvalidOperationException>(() => rendered.RequireNoLoss());
        } else rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("left-middle", "margin-left:auto;margin-right:auto", 30D)]
    [InlineData("right-middle", "margin-left:auto;margin-right:auto", 340D)]
    [InlineData("left-middle", "margin-left:5px;margin-right:auto", 5D)]
    [InlineData("right-middle", "margin-left:auto;margin-right:5px", 355D)]
    public void HtmlPageMarginSideCorner_AutomaticSideMarginsSolveTheFixedWidthEquation(
        string position, string margins, double x) {
        HtmlRenderDocument rendered = RenderSideCornerMargins("@" + position
            + "{content:'';width:40px;height:20px;background:blue;" + margins + "}");
        AssertSideCornerMarginRectangle(SideCornerMarginFill(Assert.Single(rendered.Pages), OfficeColor.Blue),
            x, 200D, 40D, 20D);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("left-top", 0D, 100D)]
    [InlineData("left-middle", 0D, 100D)]
    [InlineData("left-bottom", 0D, 100D)]
    [InlineData("right-top", 320D, 80D)]
    [InlineData("right-middle", 320D, 80D)]
    [InlineData("right-bottom", 320D, 80D)]
    public void HtmlPageMarginSideCorner_OneAutomaticSideBoxUsesTheWholeSideStrip(
        string position, double x, double width) {
        HtmlRenderDocument rendered = RenderSideCornerMargins("@" + position + "{content:'A';background:blue}");
        AssertSideCornerMarginRectangle(SideCornerMarginFill(Assert.Single(rendered.Pages), OfficeColor.Blue),
            x, 60D, width, 300D);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("left")]
    [InlineData("right")]
    public void HtmlPageMarginSideCorner_AutoHeightAllocationKeepsTheMiddleCentredBesideUnequalContent(string side) {
        HtmlRenderDocument rendered = RenderSideCornerMargins(
            "@" + side + "-top{content:'A\\A B\\A C';white-space:pre;background:red}"
            + "@" + side + "-middle{content:'M';background:blue}"
            + "@" + side + "-bottom{content:'Z';background:green}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderShape top = SideCornerMarginFill(page, OfficeColor.Red);
        HtmlRenderShape middle = SideCornerMarginFill(page, OfficeColor.Blue);
        HtmlRenderShape bottom = SideCornerMarginFill(page, OfficeColor.Green);

        // The middle's 20px contribution competes with an imaginary pair whose
        // contribution is twice the larger (60px) side, preserving its centre.
        Assert.Equal(900D / 7D, top.Height, 6);
        Assert.Equal(300D / 7D, middle.Height, 6);
        Assert.Equal(900D / 7D, bottom.Height, 6);
        Assert.Equal(210D, middle.Y + middle.Height / 2D, 6);
        Assert.Equal(60D, top.Y, 6);
        Assert.Equal(360D, bottom.Y + bottom.Height, 6);
        Assert.Equal("A\nB\nC", PageMarginTypographyText(page, side + "-top").Text);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("left")]
    [InlineData("right")]
    public void HtmlPageMarginSideCorner_FixedSideHeightLeavesTheRemainderToItsAutomaticPeer(string side) {
        HtmlRenderDocument rendered = RenderSideCornerMargins(
            "@" + side + "-top{content:'';height:80px;background:red}"
            + "@" + side + "-bottom{content:'B';background:blue}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderShape top = SideCornerMarginFill(page, OfficeColor.Red);
        HtmlRenderShape bottom = SideCornerMarginFill(page, OfficeColor.Blue);
        Assert.Equal(60D, top.Y, 6);
        Assert.Equal(80D, top.Height, 6);
        Assert.Equal(140D, bottom.Y, 6);
        Assert.Equal(220D, bottom.Height, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("left")]
    [InlineData("right")]
    public void HtmlPageMarginSideCorner_FixedHeightsKeepBothEdgesAndTheMiddleAnchor(string side) {
        HtmlRenderDocument rendered = RenderSideCornerMargins(
            "@" + side + "-top{content:'';height:60px;background:red}"
            + "@" + side + "-middle{content:'';height:40px;background:blue}"
            + "@" + side + "-bottom{content:'';height:20px;background:green}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderShape top = SideCornerMarginFill(page, OfficeColor.Red);
        HtmlRenderShape middle = SideCornerMarginFill(page, OfficeColor.Blue);
        HtmlRenderShape bottom = SideCornerMarginFill(page, OfficeColor.Green);
        Assert.Equal(60D, top.Y, 6);
        Assert.Equal(60D, top.Height, 6);
        Assert.Equal(190D, middle.Y, 6);
        Assert.Equal(40D, middle.Height, 6);
        Assert.Equal(340D, bottom.Y, 6);
        Assert.Equal(20D, bottom.Height, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("left-top", "height:80px;max-height:50px", 50D)]
    [InlineData("right-top", "height:80px;min-height:100px;max-height:60px", 100D)]
    public void HtmlPageMarginSideCorner_SideHeightConstraintsApplyMaximumThenMinimum(
        string position, string constraint, double height) {
        HtmlRenderDocument rendered = RenderSideCornerMargins("@" + position
            + "{content:'';background:blue;" + constraint + "}");
        HtmlRenderShape fill = SideCornerMarginFill(Assert.Single(rendered.Pages), OfficeColor.Blue);
        Assert.Equal(60D, fill.Y, 6);
        Assert.Equal(height, fill.Height, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("left-top", 60D)]
    [InlineData("left-middle", 200D)]
    [InlineData("left-bottom", 340D)]
    [InlineData("right-top", 60D)]
    [InlineData("right-middle", 200D)]
    [InlineData("right-bottom", 340D)]
    [InlineData("top-left-corner", 30D)]
    [InlineData("top-right-corner", 30D)]
    [InlineData("bottom-left-corner", 370D)]
    [InlineData("bottom-right-corner", 370D)]
    public void HtmlPageMarginSideCorner_DefaultVerticalAlignmentUsesTheNamedPosition(
        string position, double textY) {
        HtmlRenderDocument rendered = RenderSideCornerMargins("@" + position
            + "{content:'A';width:40px;height:40px;background:blue}");
        Assert.Equal(textY, PageMarginTypographyText(Assert.Single(rendered.Pages), position).LayoutY, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("left-top", "bottom", 80D)]
    [InlineData("right-middle", "top", 190D)]
    [InlineData("left-bottom", "middle", 330D)]
    [InlineData("top-right-corner", "top", 20D)]
    public void HtmlPageMarginSideCorner_AuthoredVerticalAlignmentOverridesThePositionDefault(
        string position, string alignment, double textY) {
        HtmlRenderDocument rendered = RenderSideCornerMargins("@" + position
            + "{content:'A';width:40px;height:40px;background:blue;vertical-align:" + alignment + "}");
        Assert.Equal(textY, PageMarginTypographyText(Assert.Single(rendered.Pages), position).LayoutY, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("inherit", 210D)]
    [InlineData("initial", 190D)]
    [InlineData("unset", 190D)]
    [InlineData("revert", 200D)]
    public void HtmlPageMarginSideCorner_VerticalAlignmentCssWideValuesUseTheCorrectContext(
        string value, double textY) {
        HtmlRenderDocument rendered = RenderSideCornerMargins("vertical-align:bottom;"
            + "@left-middle{content:'A';width:40px;height:40px;background:blue;vertical-align:" + value + "}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        Assert.Equal(190D, SideCornerMarginFill(page, OfficeColor.Blue).Y, 6);
        // Initial/unset select baseline in this independent cell, without peer displacement.
        // Revert restores the positional middle alignment.
        Assert.Equal(textY, PageMarginTypographyText(page, "left-middle").LayoutY, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("top-left-corner", "", OfficeTextAlignment.Right)]
    [InlineData("bottom-left-corner", "", OfficeTextAlignment.Right)]
    [InlineData("top-right-corner", "", OfficeTextAlignment.Left)]
    [InlineData("bottom-right-corner", "", OfficeTextAlignment.Left)]
    [InlineData("top-left-corner", "text-align:left", OfficeTextAlignment.Left)]
    [InlineData("bottom-right-corner", "text-align:right", OfficeTextAlignment.Right)]
    public void HtmlPageMarginSideCorner_CornerTextAlignmentUsesTheDefaultOrExplicitAuthorValue(
        string position, string declaration, OfficeTextAlignment expected) {
        HtmlRenderDocument rendered = RenderSideCornerMargins("@" + position
            + "{content:'A';width:40px;height:40px;background:blue;" + declaration + "}");
        HtmlRenderText text = PageMarginTypographyText(Assert.Single(rendered.Pages), position);
        Assert.Equal(expected, text.Alignment);
        Assert.Equal("A", text.Text);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("top-left-corner", 0D, 0D, 100D, 60D)]
    [InlineData("top-right-corner", 320D, 0D, 80D, 60D)]
    [InlineData("bottom-left-corner", 0D, 360D, 100D, 40D)]
    [InlineData("bottom-right-corner", 320D, 360D, 80D, 40D)]
    public void HtmlPageMarginSideCorner_AutomaticCornerDimensionsUseTheMarginIntersection(
        string position, double x, double y, double width, double height) {
        HtmlRenderDocument rendered = RenderSideCornerMargins("@" + position + "{content:'';background:blue}");
        AssertSideCornerMarginRectangle(SideCornerMarginFill(Assert.Single(rendered.Pages), OfficeColor.Blue),
            x, y, width, height);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("top-left-corner", "auto", 30D, 20D)]
    [InlineData("bottom-right-corner", "auto", 340D, 370D)]
    [InlineData("top-left-corner", "-5px", 65D, 45D)]
    [InlineData("bottom-right-corner", "-5px", 315D, 355D)]
    public void HtmlPageMarginSideCorner_CornerMarginsSolveBothFixedAxes(
        string position, string margin, double x, double y) {
        HtmlRenderDocument rendered = RenderSideCornerMargins("@" + position
            + "{content:'';width:40px;height:20px;margin:" + margin + ";background:blue}");
        AssertSideCornerMarginRectangle(SideCornerMarginFill(Assert.Single(rendered.Pages), OfficeColor.Blue),
            x, y, 40D, 20D);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("left-middle", "''", 40, 20, true)]
    [InlineData("top-left-corner", "''", 40, 20, true)]
    [InlineData("left-middle", "none", 40, 20, false)]
    [InlineData("top-left-corner", "normal", 40, 20, false)]
    [InlineData("left-middle", "''", 0, 20, false)]
    [InlineData("top-left-corner", "''", 40, 0, false)]
    public void HtmlPageMarginSideCorner_OnlyGeneratedPositiveAreaBoxesPaintTheirEmptyDecoration(
        string position, string content, int width, int height, bool painted) {
        HtmlRenderDocument rendered = RenderSideCornerMargins("@" + position + "{content:" + content
            + ";width:" + width + "px;height:" + height + "px;background:blue}");
        HtmlRenderShape[] fills = EnumerateRenderVisuals(Assert.Single(rendered.Pages).Scene).OfType<HtmlRenderShape>()
            .Where(shape => shape.Shape.FillColor == OfficeColor.Blue).ToArray();
        Assert.Equal(painted ? 1 : 0, fills.Length);
        Assert.Equal("BODY", rendered.Text.Trim());
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("left-middle")]
    [InlineData("bottom-right-corner")]
    public void HtmlPageMarginSideCorner_CountersUseFinalPagesWithoutDuplicatingBodyText(string position) {
        HtmlRenderDocument rendered = RenderSideCornerMargins("@" + position
            + "{content:counter(page) '/' counter(pages);width:40px;height:20px;background:blue}",
            "<p>FIRST</p><p style='break-before:page'>SECOND</p>");
        Assert.Equal(2, rendered.Pages.Count);
        Assert.Equal("1/2", PageMarginTypographyText(rendered.Pages[0], position).Text);
        Assert.Equal("2/2", PageMarginTypographyText(rendered.Pages[1], position).Text);
        Assert.Equal("FIRSTSECOND", PageMarginTypographyCompact(rendered.Text));
        Assert.All(rendered.Pages, page => Assert.Equal(40D, SideCornerMarginFill(page, OfficeColor.Blue).Width, 6));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("left-middle", true)]
    [InlineData("top-left-corner", true)]
    [InlineData("left-middle", false)]
    [InlineData("top-left-corner", false)]
    public void HtmlPageMarginSideCorner_ImageBackgroundsUseTheConfiguredResourcePolicy(string position, bool allowed) {
        string data = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(255, 0, 0));
        HtmlRenderOptions options = SideCornerMarginOptions();
        if (!allowed) options.UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile();
        HtmlRenderDocument rendered = RenderSideCornerMargins("@" + position
            + "{content:'';width:40px;height:20px;background:blue;background-image:url(data:image/png;base64," + data
            + ");background-size:10px 10px;background-repeat:no-repeat}", options: options);
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderShape fill = SideCornerMarginFill(page, OfficeColor.Blue);
        OfficeRasterImage pixels = OfficeDrawingRasterRenderer.Render(page.CreateDrawing());
        Assert.Equal(allowed ? OfficeColor.Red : OfficeColor.Blue, pixels.GetPixel((int)fill.X + 2, (int)fill.Y + 2));
        if (allowed) rendered.RequireNoLoss();
        else Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == "ImageResourceRejectedByPolicy");
    }

    [Theory]
    [InlineData("left-middle")]
    [InlineData("top-left-corner")]
    public void HtmlPageMarginSideCorner_ImageTilingRetainsTheOperationBudgetAndStrictLoss(string position) {
        string data = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(255, 0, 0));
        HtmlRenderOptions options = SideCornerMarginOptions();
        options.MaxBackgroundImageTiles = 8;
        HtmlRenderDocument rendered = RenderSideCornerMargins("@" + position
            + "{content:'';width:40px;height:20px;background-image:url(data:image/png;base64," + data
            + ");background-size:1px 1px}", options: options);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageTileLimitExceeded
            && diagnostic.Source == "@page @" + position && diagnostic.LossKind != OfficeConversionLossKind.None);
        Assert.ThrowsAny<InvalidOperationException>(() => rendered.RequireNoLoss());
    }

    [Theory]
    [InlineData("height:max-content")]
    [InlineData("height:0;min-height:min-content")]
    [InlineData("height:80px;max-height:max-content")]
    public void HtmlPageMarginSideCorner_ZeroSizeTextRetainsIntrinsicLeadingWithoutGlyphPaint(string constraint) {
        HtmlRenderDocument rendered = RenderSideCornerMargins(
            "@left-top{content:'A\\A B';white-space:pre;font-size:0;line-height:20px;width:40px;"
            + constraint + ";padding:3px;border:2px solid green;background:blue}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        AssertSideCornerMarginRectangle(SideCornerMarginFill(page, OfficeColor.Blue), 50D, 60D, 50D, 50D);
        Assert.DoesNotContain(EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderText>(),
            text => text.Source == "@page @left-top");
        Assert.Equal("BODY", rendered.Text.Trim());
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginSideCorner_ZeroSizeTextContributesLeadingToAutomaticPeerHeights() {
        HtmlRenderDocument rendered = RenderSideCornerMargins(
            "@left-top{content:'A\\A B';white-space:pre;font-size:0;line-height:20px;background:red}"
            + "@left-middle{content:'M';background:blue}@left-bottom{content:'B';background:green}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        Assert.Equal(120D, SideCornerMarginFill(page, OfficeColor.Red).Height, 6);
        Assert.Equal(60D, SideCornerMarginFill(page, OfficeColor.Blue).Height, 6);
        Assert.Equal(120D, SideCornerMarginFill(page, OfficeColor.Green).Height, 6);
        Assert.DoesNotContain(EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderText>(), text => text.Source == "@page @left-top");
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("left", 70, 100)]
    [InlineData("bottom", 120, 370)]
    public void HtmlPageMarginSideCorner_OverlappingBoxesUseClockwiseDefaultPaintOrder(string pair, int x, int y) {
        string boxes = pair == "left"
            ? "@left-top{content:'';width:40px;height:200px;background:red}"
                + "@left-middle{content:'';width:40px;height:300px;background:blue}"
            : "@bottom-left-corner{content:'';width:150px;height:40px;margin-right:-60px;background:red}"
                + "@bottom-left{content:'';height:40px;background:blue}";
        HtmlRenderDocument rendered = RenderSideCornerMargins(boxes);
        OfficeRasterImage pixels = OfficeDrawingRasterRenderer.Render(Assert.Single(rendered.Pages).CreateDrawing());
        Assert.Equal(OfficeColor.Red, pixels.GetPixel(x, y));
        Assert.Equal("BODY", rendered.Text.Trim());
        rendered.RequireNoLoss();
    }

    private static HtmlRenderDocument RenderSideCornerMargins(string boxes, string body = "<p>BODY</p>",
        HtmlRenderOptions? options = null) {
        string family = ManagedTextShapingTestAssets.FamilyName;
        string html = "<style>html,body,p{margin:0;padding:0;font:20px/20px '" + family
            + "';font-kerning:none}@page{size:400px 400px;margin:60px 80px 40px 100px;font:20px/20px '"
            + family + "';" + boxes + "}</style>" + body;
        return HtmlRenderTestDriver.Render(html, options ?? SideCornerMarginOptions());
    }

    private static HtmlRenderOptions SideCornerMarginOptions() {
        string family = ManagedTextShapingTestAssets.FamilyName;
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, HonorCssPageRules = true,
            AllowSystemFontFallback = false, DefaultFontFamily = family,
            UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser
        };
        options.Fonts.Add(family, ManagedTextShapingTestAssets.CreateFontWithScalarAdvances(
            Enumerable.Range(32, 95).ToDictionary(scalar => scalar, _ => (ushort)500)));
        return options;
    }

    private static HtmlRenderShape SideCornerMarginFill(HtmlRenderPage page, OfficeColor color) =>
        Assert.Single(EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderShape>(), shape => shape.Shape.FillColor == color);

    private static void AssertSideCornerMarginRectangle(HtmlRenderVisual paint, double x, double y, double width, double height) {
        Assert.Equal(x, paint.X, 6);
        Assert.Equal(y, paint.Y, 6);
        Assert.Equal(width, paint.Width, 6);
        Assert.Equal(height, paint.Height, 6);
    }
}

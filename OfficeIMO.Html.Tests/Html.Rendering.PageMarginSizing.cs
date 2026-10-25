using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("top", true, true)]
    [InlineData("bottom", true, true)]
    [InlineData("top", true, false)]
    [InlineData("top", false, true)]
    public void HtmlPageMarginSizing_CentreUsesSpareWidthWithoutMovingOffCentre(string side, bool left, bool right) {
        string boxes = "@" + side + "-center{content:'TITLE CENTRE'}"
            + (left ? "@" + side + "-left{content:'LL'}" : "")
            + (right ? "@" + side + "-right{content:'RR'}" : "");
        HtmlRenderDocument rendered = RenderMarginSizing(300D, boxes);
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderText centre = PageMarginTypographyText(page, side + "-center");

        Assert.Equal(225D, centre.Width, 6);
        Assert.Equal(page.Margins.Left + 150D, centre.X + centre.Width / 2D, 6);
        Assert.Single(MarginSizingPaintedText(page, centre));
        Assert.Equal("TITLE CENTRE", centre.Text);
        if (left) Assert.Equal(37.5D, PageMarginTypographyText(page, side + "-left").Width, 6);
        if (right) Assert.Equal(37.5D, PageMarginTypographyText(page, side + "-right").Width, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("top", 300D, "AAA", "R", 225D, 75D)]
    [InlineData("bottom", 100D, "AAAA AA", "RR RR", 60D, 40D)]
    [InlineData("top", 50D, "AAAA", "RR", 100D / 3D, 50D / 3D)]
    public void HtmlPageMarginSizing_EdgesShareSpaceByTheirContentAcrossConstraintRegimes(
        string side, double availableWidth, string leftText, string rightText, double leftWidth, double rightWidth) {
        string boxes = "@" + side + "-left{content:'" + leftText + "'}@" + side + "-right{content:'" + rightText + "'}";
        HtmlRenderDocument rendered = RenderMarginSizing(availableWidth, boxes);
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderText left = PageMarginTypographyText(page, side + "-left");
        HtmlRenderText right = PageMarginTypographyText(page, side + "-right");

        Assert.Equal(leftWidth, left.Width, 6);
        Assert.Equal(rightWidth, right.Width, 6);
        Assert.Equal(page.Margins.Left, left.X, 6);
        Assert.Equal(page.Width - page.Margins.Right, right.X + right.Width, 6);
        Assert.Equal(leftText, left.Text);
        Assert.Equal(rightText, right.Text);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("AA AA", "BBBB BBBB", "RR", 160D, 840D / 11D)]
    [InlineData("AAAA", "BBBBBBBB", "R", 80D, 40D)]
    public void HtmlPageMarginSizing_ConstrainedCentreKeepsCompleteTextInItsAllocatedFrame(
        string left, string centreText, string right, double availableWidth, double expectedCentreWidth) {
        string boxes = "@top-left{content:'" + left + "'}@top-center{content:'" + centreText
            + "'}@top-right{content:'" + right + "'}";
        HtmlRenderDocument rendered = RenderMarginSizing(availableWidth, boxes);
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderText centre = PageMarginTypographyText(page, "top-center");

        Assert.Equal(expectedCentreWidth, centre.Width, 6);
        Assert.Equal(page.Margins.Left + availableWidth / 2D, centre.X + centre.Width / 2D, 6);
        Assert.Equal(centreText, centre.Text);
        OfficeDrawingText[] lines = MarginSizingPaintedText(page, centre);
        Assert.True(lines.Length > 1);
        Assert.Equal(PageMarginTypographyCompact(centreText),
            PageMarginTypographyCompact(string.Concat(lines.Select(line => line.Text))));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("none", 225D)]
    [InlineData("normal", 225D)]
    [InlineData("''", 150D)]
    public void HtmlPageMarginSizing_GeneratedEmptyCentreDiffersFromAnAbsentCentre(string content, double leftWidth) {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-left{content:'AAA'}@top-center{content:" + content + "}@top-right{content:'R'}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);

        Assert.Equal(leftWidth, PageMarginTypographyText(page, "top-left").Width, 6);
        Assert.DoesNotContain(EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderText>(),
            text => text.Source == "@page @top-center");
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginSizing_ZeroAdvanceContentUsesFiniteEqualFactors() {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-left{content:'L'}@top-center{content:'C'}@top-right{content:'R'}", zeroAdvance: true);
        HtmlRenderPage page = Assert.Single(rendered.Pages);

        Assert.Equal(75D, PageMarginTypographyText(page, "top-left").Width, 6);
        Assert.Equal(150D, PageMarginTypographyText(page, "top-center").Width, 6);
        Assert.Equal(75D, PageMarginTypographyText(page, "top-right").Width, 6);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginSizing_TabExpansionAndHardLinesMatchThePaintedTextBlock() {
        HtmlRenderDocument rendered = RenderMarginSizing(180D,
            "@top-left{content:'L'}@top-center{content:'A\\9 B\\A C'}@top-right{content:'R'}");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderText centre = PageMarginTypographyText(page, "top-center");

        Assert.Equal(900D / 7D, centre.Width, 6);
        OfficeDrawingText[] lines = MarginSizingPaintedText(page, centre);
        Assert.Equal(new[] { "A   B", "C" }, lines.Select(line => line.Text).ToArray());
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginSizing_PageSpecificCountersAreMeasuredAfterPagination() {
        string body = string.Join("", Enumerable.Range(1, 10).Select(page =>
            "<p" + (page == 1 ? "" : " style='break-before:page'") + ">BODY</p>"));
        HtmlRenderDocument rendered = RenderMarginSizing(120D,
            "@top-left{content:'L'}@top-center{content:counter(page)}@top-right{content:'R'}", body: body);

        Assert.Equal(10, rendered.Pages.Count);
        Assert.Equal(40D, PageMarginTypographyText(rendered.Pages[0], "top-center").Width, 6);
        Assert.Equal(60D, PageMarginTypographyText(rendered.Pages[9], "top-center").Width, 6);
        Assert.Equal("10", PageMarginTypographyText(rendered.Pages[9], "top-center").Text);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginSizing_AnEmptyFirstExceptStringStillGeneratesTheCentre() {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-left{content:'AAA'}@top-center{content:string(chapter,first-except)}@top-right{content:'R'}",
            body: "<p style='string-set:chapter content()'>TITLE</p><p style='break-before:page'>BODY</p>");

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Equal(150D, PageMarginTypographyText(rendered.Pages[0], "top-left").Width, 6);
        Assert.Equal(1500D / 11D, PageMarginTypographyText(rendered.Pages[1], "top-center").Width, 6);
        Assert.Equal("TITLE", PageMarginTypographyText(rendered.Pages[1], "top-center").Text);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlPageMarginSizing_RunningElementUsesSourceIntrinsicsAndTheSharedAssignedWidth() {
        HtmlRenderDocument rendered = RenderMarginSizing(300D,
            "@top-left{content:element(heading)}@top-center{content:'TITLE CENTRE'}@top-right{content:'R'}",
            body: "<p id='running' style='position:running(heading)'>LEFT</p><p>BODY</p>");
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup running = Assert.Single(page.Scene.OfType<HtmlRenderSemanticGroup>(),
            group => group.Source == "@page @top-left element(heading)");

        Assert.Equal(60D, running.Width, 6);
        Assert.Equal(180D, PageMarginTypographyText(page, "top-center").Width, 6);
        HtmlRenderText text = Assert.Single(EnumerateRenderVisuals(running.Visuals).OfType<HtmlRenderText>());
        Assert.Equal("LEFT", text.Text);
        Assert.Equal(20D, text.Font.Size, 6);
        Assert.InRange(text.X + text.Width, running.X, running.X + running.Width);
        Assert.Equal("BODY", rendered.Text.Trim());
        rendered.RequireNoLoss();
    }

    private static OfficeDrawingText[] MarginSizingPaintedText(HtmlRenderPage page, HtmlRenderText margin) =>
        PageMarginTypographyDrawingText(page.CreateDrawing())
            .Where(text => (margin.Source!.StartsWith("@page @top-", StringComparison.Ordinal)
                    ? text.Y < page.Margins.Top : text.Y > page.Height - page.Margins.Bottom)
                && text.X >= margin.X - 0.001D && text.X < margin.X + margin.Width - 0.001D).ToArray();

    private static HtmlRenderDocument RenderMarginSizing(double availableWidth, string boxes,
        string body = "<p>BODY</p>", bool zeroAdvance = false, ushort zeroGlyphAdvance = 500) {
        string family = ManagedTextShapingTestAssets.FamilyName;
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, HonorCssPageRules = true,
            AllowSystemFontFallback = false, DefaultFontFamily = family,
            UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser
        };
        options.Fonts.Add(family, ManagedTextShapingTestAssets.CreateFontWithScalarAdvances(
            Enumerable.Range(32, 95).ToDictionary(scalar => scalar,
                scalar => zeroAdvance ? (ushort)0 : scalar == '0' ? zeroGlyphAdvance : (ushort)500)));
        string width = (availableWidth + 40D).ToString("R", System.Globalization.CultureInfo.InvariantCulture);
        string html = "<style>html,body,p{margin:0;padding:0;font:20px/20px '" + family
            + "';font-kerning:none}@page{size:" + width + "px 300px;margin:60px 20px;font:20px/20px '"
            + family + "';" + boxes + "}</style>" + body;
        return HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options);
    }
}

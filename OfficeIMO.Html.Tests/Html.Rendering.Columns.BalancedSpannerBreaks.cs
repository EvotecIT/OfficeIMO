using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("auto")]
    [InlineData("balance")]
    public void HtmlColumns_EmptySpannerPageDoesNotPaintTheFollowingCollapsedTable(string fill) {
        HtmlRenderDocument rendered = RenderSpannerBreakControls(HtmlRenderIntentProfile.PrintPaged,
            "<p>P1</p><p>P2</p><div style='column-span:all;break-before:page;break-after:page'></div>"
            + "<table id='empty-page-tail' style='column-span:all;border-collapse:collapse;width:200px'>"
            + "<tr><td style='border:2px solid red;padding:0'>TABLE</td></tr></table><p>T1</p>",
            extraStyle: "column-fill:" + fill);

        Assert.Equal(3, rendered.Pages.Count);
        Assert.Empty(ColumnPageBandText(rendered.Pages[1]));
        Assert.DoesNotContain(EnumerateRenderVisuals(rendered.Pages[1].Scene).OfType<HtmlRenderShape>(),
            shape => shape.Source?.StartsWith("table#empty-page-tail:collapsed-border-", StringComparison.Ordinal) == true);
        Assert.Contains(EnumerateRenderVisuals(rendered.Pages[2].Scene).OfType<HtmlRenderShape>(),
            shape => shape.Source?.StartsWith("table#empty-page-tail:collapsed-border-", StringComparison.Ordinal) == true);
        AssertColumnPageBandTextOnce(rendered, new[] { "P1", "P2", "TABLE", "T1", "AFTER" });
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, true)]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, false)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, true)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, false)]
    public void HtmlColumns_BalancedSpannerPreservesItsAuthoredPageForce(HtmlRenderIntentProfile profile, bool before) {
        HtmlRenderDocument rendered = RenderBalancedSpannerBreakControls(profile,
            "<p>P1</p><p>P2</p><p style='column-span:all;break-" + (before ? "before" : "after")
            + ":page'>SPAN</p><p>T1</p><p>T2</p>");

        Assert.Equal(2, rendered.Pages.Count);
        AssertColumnPageBandTextOnce(rendered, new[] { "P1", "P2", "SPAN", "T1", "T2", "AFTER" });
        Assert.Equal(before ? 2 : 1, ColumnPageBandTextPage(rendered, "SPAN").PageNumber);
        Assert.Equal(2, ColumnPageBandTextPage(rendered, "T1").PageNumber);
        AssertColumnBreakPlacements(rendered, ("P1", 20D, 20D), ("P2", 180D, 20D),
            ("SPAN", 20D, before ? 20D : 40D), ("T1", 20D, before ? 40D : 20D),
            ("T2", 180D, before ? 40D : 20D), ("AFTER", 20D, before ? 60D : 40D));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, true)]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, false)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, true)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, false)]
    public void HtmlColumns_BalancedSpannerInternalPageForceOverridesItsKeep(HtmlRenderIntentProfile profile, bool before) {
        HtmlRenderDocument rendered = RenderBalancedSpannerBreakControls(profile,
            "<p>P1</p><p>P2</p><div style='column-span:all;break-inside:avoid-page'>"
            + "<p>S1</p><p" + (before ? "" : " style='break-after:page'") + ">S2</p>"
            + "<p" + (before ? " style='break-before:page'" : "") + ">S3</p><p>S4</p></div>"
            + "<p>T1</p><p>T2</p>");

        Assert.Equal(2, rendered.Pages.Count);
        AssertColumnPageBandTextOnce(rendered, new[] { "P1", "P2", "S1", "S2", "S3", "S4", "T1", "T2", "AFTER" });
        Assert.Equal(1, ColumnPageBandTextPage(rendered, "S2").PageNumber);
        Assert.Equal(2, ColumnPageBandTextPage(rendered, "S3").PageNumber);
        AssertColumnBreakPlacements(rendered, ("S1", 20D, 40D), ("S2", 20D, 60D),
            ("S3", 20D, 20D), ("S4", 20D, 40D), ("T1", 20D, 60D),
            ("T2", 180D, 60D), ("AFTER", 20D, 80D));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, true)]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, false)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, true)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, false)]
    public void HtmlColumns_BalancedSpannerPageKeepMovesItsFittingNeighbour(HtmlRenderIntentProfile profile, bool before) {
        HtmlRenderDocument rendered = RenderBalancedSpannerBreakControls(profile,
            "<p>P1</p><p>P2</p><p style='column-span:all;break-" + (before ? "before" : "after")
            + ":avoid-page'>SPAN</p><p>T1</p>",
            "<p style='height:" + (before ? "80" : "60") + "px'>ENTRY</p>");

        Assert.Equal(2, rendered.Pages.Count);
        AssertColumnPageBandTextOnce(rendered, new[] { "ENTRY", "P1", "P2", "SPAN", "T1", "AFTER" });
        Assert.Equal(before ? 2 : 1, ColumnPageBandTextPage(rendered, "P1").PageNumber);
        Assert.Equal(2, ColumnPageBandTextPage(rendered, "SPAN").PageNumber);
        Assert.Equal(2, ColumnPageBandTextPage(rendered, "T1").PageNumber);
        AssertColumnBreakPlacements(rendered, ("SPAN", 20D, before ? 40D : 20D),
            ("T1", 20D, before ? 60D : 40D), ("AFTER", 20D, before ? 80D : 60D));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, false, 2)]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, true, 3)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, false, 2)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, true, 3)]
    public void HtmlColumns_BalancedSpannersPreserveSharedAndDistinctEmptyBoundaries(
        HtmlRenderIntentProfile profile, bool empty, int pages) {
        string first = empty ? "<div style='column-span:all;break-before:page;break-after:page'></div>"
            : "<p style='column-span:all;break-after:page'>S1</p>";
        HtmlRenderDocument rendered = RenderBalancedSpannerBreakControls(profile,
            "<p>P1</p><p>P2</p>" + first + "<p style='column-span:all;break-before:page'>S2</p><p>T1</p><p>T2</p>");

        Assert.Equal(pages, rendered.Pages.Count);
        Assert.Equal(pages, ColumnPageBandTextPage(rendered, "S2").PageNumber);
        AssertColumnPageBandTextOnce(rendered, empty ? new[] { "P1", "P2", "S2", "T1", "T2", "AFTER" }
            : new[] { "P1", "P2", "S1", "S2", "T1", "T2", "AFTER" });
        if (empty) Assert.Empty(ColumnPageBandText(rendered.Pages[1]));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, true)]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, false)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, true)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, false)]
    public void HtmlColumns_BalancedSpannerPageForcesRetainAdjacentColumnForces(HtmlRenderIntentProfile profile, bool before) {
        HtmlRenderDocument rendered = RenderBalancedSpannerBreakControls(profile,
            "<p" + (before ? " style='break-after:column'" : "") + ">P1</p><p>P2</p>"
            + "<p style='column-span:all;break-" + (before ? "before" : "after") + ":page'>SPAN</p>"
            + "<p" + (before ? "" : " style='break-after:column'") + ">T1</p><p>T2</p>");

        Assert.Equal(2, rendered.Pages.Count);
        AssertColumnPageBandTextOnce(rendered, new[] { "P1", "P2", "SPAN", "T1", "T2", "AFTER" });
        AssertColumnBreakPlacements(rendered, ("P1", 20D, 20D), ("P2", 180D, 20D),
            ("T1", 20D, before ? 40D : 20D), ("T2", 180D, before ? 40D : 20D));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("transform:translateY(0)")]
    [InlineData("overflow:hidden")]
    public void HtmlColumns_IneligibleBalancedSpannerHostsRetainTypedLoss(string hostStyle) {
        HtmlRenderDocument rendered = RenderSpannerBreakControls(HtmlRenderIntentProfile.PrintPaged,
            "<p>P1</p><p>P2</p><p style='column-span:all;break-before:page'>SPAN</p><p>T1</p>",
            extraStyle: "column-fill:balance;" + hostStyle);

        HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>().ToArray();
        foreach (string label in new[] { "P1", "P2", "SPAN", "T1", "AFTER" })
            Assert.Single(text, item => item.Text == label);
        Assert.Contains(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.MultiColumnValueUnsupported);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

    [Fact]
    public void HtmlColumns_BalancedSpannerPageForceObservesTheMaximumPageCount() {
        Assert.Throws<InvalidOperationException>(() => RenderSpannerBreakControls(HtmlRenderIntentProfile.PrintPaged,
            "<p>P1</p><p>P2</p><p style='column-span:all;break-before:page'>SPAN</p><p>T1</p>",
            extraStyle: "column-fill:balance", maxPageCount: 1));
    }

    private static HtmlRenderDocument RenderBalancedSpannerBreakControls(
        HtmlRenderIntentProfile profile, string body, string entry = "") =>
        RenderSpannerBreakControls(profile, body, entry, "column-fill:balance");
}

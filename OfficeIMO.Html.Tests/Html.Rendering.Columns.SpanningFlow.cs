using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("column-fill:auto;max-height:100px", 0D, 80D, 160D, 40D, 100D)]
    [InlineData("column-fill:balance;max-height:100px", 160D, 40D, 160D, 60D, 80D)]
    [InlineData("column-fill:auto;min-height:100px", 0D, 80D, 0D, 100D, 120D)]
    public void HtmlColumns_SpanningTerminalDistinguishesMaximumAndMinimumHeight(
        string style, double thirdX, double thirdY, double fourthX, double fourthY, double afterY) {
        HtmlRenderDocument rendered = RenderColumnBreaks(
            "<p>P1</p><p>P2</p><p style='column-span:all'>SPAN</p>"
            + "<p>T1</p><p>T2</p><p>T3</p><p>T4</p>", style);

        AssertColumnBreakPlacements(rendered, ("SPAN", 0D, 20D), ("T1", 0D, 40D),
            ("T2", 0D, 60D), ("T3", thirdX, thirdY), ("T4", fourthX, fourthY),
            ("BREAK-AFTER", 0D, afterY));
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged)]
    public void HtmlColumns_SpanningPaddedHeightRetainsBlankTailAcrossPhysicalPages(HtmlRenderIntentProfile profile) {
        string html = ColumnBreakSource("<style>@page{size:700px 140px;margin:20px}</style>"
            + "<section style='width:300px;box-sizing:border-box;height:120px;padding:10px;"
            + "column-count:2;column-gap:20px;column-fill:auto'>"
            + "<p>P1</p><p>P2</p><p style='column-span:all'>SPAN</p>"
            + "<p>T1</p><p>T2</p><p>T3</p><p>T4</p></section><p>AFTER</p>");
        HtmlRenderDocument rendered = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
            HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, ColumnPageBandOptions())).Document;

        Assert.Equal(2, rendered.Pages.Count);
        AssertColumnPageBandTextOnce(rendered, new[] { "P1", "P2", "SPAN", "T1", "T2", "T3", "T4", "AFTER" });
        Assert.Equal(2, ColumnPageBandTextPage(rendered, "AFTER").PageNumber);
        AssertColumnBreakPlacements(rendered, ("T1", 30D, 70D), ("T2", 30D, 90D),
            ("T3", 180D, 70D), ("T4", 180D, 90D), ("AFTER", 20D, 40D));
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData("column-fill:auto", 0D, 80D, 100D)]
    [InlineData("column-fill:balance", 160D, 60D, 80D)]
    [InlineData("column-fill:balance;height:100px", 160D, 60D, 100D)]
    public void HtmlColumns_SpanningFlowBalancesBeforeTheSpannerAndAppliesTerminalFill(
        string style, double secondTailX, double secondTailY, double afterY) {
        HtmlRenderDocument rendered = RenderColumnBreaks(
            "<p>P1</p><p>P2</p><p>P3</p><p>P4</p>"
            + "<p style='column-span:all'>SPAN</p><p>T1</p><p>T2</p>", style);

        AssertColumnBreakPlacements(rendered,
            ("P1", 0D, 0D), ("P2", 0D, 20D), ("P3", 160D, 0D), ("P4", 160D, 20D),
            ("SPAN", 0D, 40D), ("T1", 0D, 60D), ("T2", secondTailX, secondTailY),
            ("BREAK-AFTER", 0D, afterY));
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData(false, 0D, 100D, 120D)]
    [InlineData(true, 160D, 60D, 100D)]
    public void HtmlColumns_SpanningAutoTerminalUsesOnlyTheRemainingDeclaredHeight(
        bool definite, double thirdTailX, double thirdTailY, double afterY) {
        HtmlRenderDocument rendered = RenderColumnBreaks(
            "<p>P1</p><p>P2</p><p>P3</p><p>P4</p>"
            + "<p style='column-span:all'>SPAN</p><p>T1</p><p>T2</p><p>T3</p>",
            "column-fill:auto;" + (definite ? "height:100px" : string.Empty));

        AssertColumnBreakPlacements(rendered, ("SPAN", 0D, 40D), ("T1", 0D, 60D),
            ("T2", 0D, 80D), ("T3", thirdTailX, thirdTailY), ("BREAK-AFTER", 0D, afterY));
        AssertColumnBreaksQualified(rendered);
    }

    [Fact]
    public void HtmlColumns_SpanningPrecedingSetHonorsAnAuthoredColumnBoundary() {
        HtmlRenderDocument rendered = RenderColumnBreaks(
            "<p>P1</p><p style='break-before:column'>P2</p><p>P3</p>"
            + "<p style='column-span:all'>SPAN</p><p>T1</p>");

        AssertColumnBreakPlacements(rendered, ("P1", 0D, 0D), ("P2", 160D, 0D),
            ("P3", 160D, 20D), ("SPAN", 0D, 40D), ("T1", 0D, 60D), ("BREAK-AFTER", 0D, 80D));
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlColumns_SpanningTerminalCoalescesAdjacentForcedColumnValues(bool beforeAndAfter) {
        HtmlRenderDocument rendered = RenderColumnBreaks(
            "<p>P1</p><p>P2</p><p style='column-span:all'>SPAN</p>"
            + "<p style='break-after:column'>T1</p><p style='"
            + (beforeAndAfter ? "break-before:column" : string.Empty) + "'>T2</p><p>T3</p>");

        AssertColumnBreakPlacements(rendered, ("P1", 0D, 0D), ("P2", 160D, 0D),
            ("SPAN", 0D, 20D), ("T1", 0D, 40D), ("T2", 160D, 40D), ("T3", 160D, 60D),
            ("BREAK-AFTER", 0D, 80D));
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData("pre-trailing")]
    [InlineData("terminal-leading")]
    [InlineData("owner-edges")]
    public void HtmlColumns_SpanningPartitionEdgesDoNotCreateEmptyColumns(string boundary) {
        HtmlRenderDocument rendered = RenderColumnBreaks(
            "<p style='" + (boundary == "owner-edges" ? "break-before:column" : string.Empty) + "'>P1</p>"
            + "<p style='" + (boundary == "pre-trailing" ? "break-after:column" : string.Empty) + "'>P2</p>"
            + "<p style='column-span:all'>SPAN</p>"
            + "<p style='" + (boundary == "terminal-leading" ? "break-before:column" : string.Empty) + "'>T1</p>"
            + "<p style='" + (boundary == "owner-edges" ? "break-after:column" : string.Empty) + "'>T2</p>");

        AssertColumnBreakPlacements(rendered, ("P1", 0D, 0D), ("P2", 160D, 0D),
            ("SPAN", 0D, 20D), ("T1", 0D, 40D), ("T2", 0D, 60D), ("BREAK-AFTER", 0D, 80D));
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlColumns_SpanningTerminalKeepsBoundaryContentWithinRemainingCapacity(bool before) {
        HtmlRenderDocument rendered = RenderColumnBreaks(
            "<p style='column-span:all'>SPAN</p><p>T1</p><p>T2</p><p>T3</p>"
            + "<p style='" + (before ? string.Empty : "break-after:avoid-column") + "'>T4</p>"
            + "<p style='" + (before ? "break-before:avoid-column" : string.Empty) + "'>T5</p>",
            "column-fill:auto;height:100px");

        AssertColumnBreakPlacements(rendered, ("SPAN", 0D, 0D), ("T1", 0D, 20D),
            ("T2", 0D, 40D), ("T3", 0D, 60D), ("T4", 160D, 20D), ("T5", 160D, 40D),
            ("BREAK-AFTER", 0D, 100D));
        AssertColumnBreaksQualified(rendered);
    }

    [Fact]
    public void HtmlColumns_SpanningPrecedingSetBalancesWithoutSplittingAFittingColumnKeep() {
        HtmlRenderDocument rendered = RenderColumnBreaks(
            "<p>P1</p><p>P2</p><div style='break-inside:avoid-column'><p>G1</p><p>G2</p></div>"
            + "<p>Q1</p><p style='column-span:all'>SPAN</p><p>T1</p>");

        AssertColumnBreakPlacements(rendered, ("P1", 0D, 0D), ("P2", 0D, 20D),
            ("G1", 160D, 0D), ("G2", 160D, 20D), ("Q1", 160D, 40D),
            ("SPAN", 0D, 60D), ("T1", 0D, 80D), ("BREAK-AFTER", 0D, 100D));
        AssertColumnBreaksQualified(rendered);
    }

    [Fact]
    public void HtmlColumns_SpanningTerminalForcedBoundaryOverridesItsContainingKeep() {
        HtmlRenderDocument rendered = RenderColumnBreaks(
            "<p>P1</p><p style='column-span:all'>SPAN</p><div style='break-inside:avoid-column'>"
            + "<p>T1</p><p style='break-before:column'>T2</p><p>T3</p></div>");

        AssertColumnBreakPlacements(rendered, ("P1", 0D, 0D), ("SPAN", 0D, 20D),
            ("T1", 0D, 40D), ("T2", 160D, 40D), ("T3", 160D, 60D), ("BREAK-AFTER", 0D, 80D));
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlColumns_SpanningSetsPreserveNestedBoundaryPropagation(bool terminal) {
        string body = terminal
            ? "<p style='column-span:all'>SPAN</p><div><p>N1</p><p style='break-after:column'>N2</p></div><p>T1</p>"
            : "<p>P1</p><div><p style='break-before:column'>N1</p><p>N2</p></div>"
                + "<p style='column-span:all'>SPAN</p><p>T1</p>";

        HtmlRenderDocument rendered = RenderColumnBreaks(body);

        if (terminal) {
            AssertColumnBreakPlacements(rendered, ("SPAN", 0D, 0D), ("N1", 0D, 20D),
                ("N2", 0D, 40D), ("T1", 160D, 20D), ("BREAK-AFTER", 0D, 60D));
        } else {
            AssertColumnBreakPlacements(rendered, ("P1", 0D, 0D), ("N1", 160D, 0D),
                ("N2", 160D, 20D), ("SPAN", 0D, 40D), ("T1", 0D, 60D), ("BREAK-AFTER", 0D, 80D));
        }
        AssertColumnBreaksQualified(rendered);
    }

    [Fact]
    public void HtmlColumns_SpanningMiddleSetRetainsItsOwnForcedColumnPlan() {
        HtmlRenderDocument rendered = RenderColumnBreaks(
            "<p>P1</p><p>P2</p><p style='column-span:all'>S1</p>"
            + "<p>M1</p><p style='break-before:column'>M2</p><p>M3</p>"
            + "<p style='column-span:all'>S2</p><p>T1</p>");

        AssertColumnBreakPlacements(rendered, ("P1", 0D, 0D), ("P2", 160D, 0D),
            ("S1", 0D, 20D), ("M1", 0D, 40D), ("M2", 160D, 40D), ("M3", 160D, 60D),
            ("S2", 0D, 80D), ("T1", 0D, 100D), ("BREAK-AFTER", 0D, 120D));
        AssertColumnBreaksQualified(rendered);
    }

    [Fact]
    public void HtmlColumns_LeadingAdjacentAndTrailingSpannersHaveNoPhantomPartitionAdvance() {
        HtmlRenderDocument rendered = RenderColumnBreaks(
            "<p style='column-span:all'>S1</p><p style='column-span:all'>S2</p>"
            + "<p>M1</p><p>M2</p><p style='column-span:all'>S3</p>");

        AssertColumnBreakPlacements(rendered, ("S1", 0D, 0D), ("S2", 0D, 20D),
            ("M1", 0D, 40D), ("M2", 160D, 40D), ("S3", 0D, 60D), ("BREAK-AFTER", 0D, 80D));
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged)]
    public void HtmlColumns_SpanningAutoTerminalConsumesTheRemainingPhysicalBandAfterAColumnAdvance(
        HtmlRenderIntentProfile profile) {
        string html = ColumnBreakSource("<style>@page{size:700px 140px;margin:20px}</style>"
            + "<section style='width:300px;column-count:2;column-gap:20px;column-fill:auto'>"
            + "<p>P1</p><p>P2</p><p style='column-span:all'>SPAN</p>"
            + "<p style='break-after:column'>T1</p><p>T2</p><p>T3</p></section><p>AFTER</p>");
        HtmlRenderDocument rendered = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
            HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, ColumnPageBandOptions())).Document;

        Assert.Equal(2, rendered.Pages.Count);
        AssertColumnPageBandTextOnce(rendered, new[] { "P1", "P2", "SPAN", "T1", "T2", "T3", "AFTER" });
        foreach (string label in new[] { "P1", "P2", "SPAN", "T1", "T2", "T3" })
            Assert.Equal(1, ColumnPageBandTextPage(rendered, label).PageNumber);
        Assert.Equal(2, ColumnPageBandTextPage(rendered, "AFTER").PageNumber);
        AssertColumnBreakPlacements(rendered, ("P1", 20D, 20D), ("P2", 180D, 20D),
            ("SPAN", 20D, 40D), ("T1", 20D, 60D), ("T2", 180D, 60D), ("T3", 180D, 80D),
            ("AFTER", 20D, 20D));
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData("before")]
    [InlineData("after")]
    [InlineData("nested")]
    [InlineData("before-column")]
    [InlineData("after-column")]
    [InlineData("nested-column")]
    [InlineData("before-avoid-page")]
    public void HtmlColumns_SpannerBreaksPreserveContentAndReportTheUnqualifiedCombination(string boundary) {
        bool nested = boundary.StartsWith("nested", StringComparison.Ordinal);
        string value = boundary.EndsWith("column", StringComparison.Ordinal) ? "column"
            : boundary.EndsWith("avoid-page", StringComparison.Ordinal) ? "avoid-page" : "page";
        string spanner = nested
            ? "<div style='column-span:all'><p>SPAN</p><p style='break-before:" + value + "'>NESTED</p></div>"
            : "<p style='column-span:all;break-" + (boundary.StartsWith("after", StringComparison.Ordinal) ? "after" : "before")
                + ":" + value + "'>SPAN</p>";
        HtmlRenderDocument rendered = RenderColumnBreaks("<p>P1</p>" + spanner + "<p>T1</p>");

        AssertColumnPageBandTextOnce(rendered, new[] { "P1", "SPAN", "T1", "BREAK-AFTER" });
        if (nested) AssertColumnPageBandTextOnce(rendered, new[] { "NESTED" });
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.MultiColumnValueUnsupported);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }
}

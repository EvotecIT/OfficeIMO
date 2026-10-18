using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(8, HtmlRenderIntentProfile.PrintPaged)]
    [InlineData(12, HtmlRenderIntentProfile.PrintPaged)]
    [InlineData(8, HtmlRenderIntentProfile.ScreenMediaPaged)]
    [InlineData(12, HtmlRenderIntentProfile.ScreenMediaPaged)]
    public void HtmlColumns_DeclaredSpanningHeightKeepsExcessColumnsInlineInItsFinalPageBand(
        int tailCount, HtmlRenderIntentProfile profile) {
        HtmlRenderDocument rendered = RenderFiniteSpanningColumns(tailCount, "height:100px", profile);

        Assert.Equal(2, rendered.Pages.Count);
        AssertFiniteSpanningTextOnce(rendered, tailCount);
        for (int index = 1; index <= tailCount; index++) {
            string label = FiniteSpanningTailLabel(index);
            Assert.Equal(1, ColumnPageBandTextPage(rendered, label).PageNumber);
            AssertColumnBreakPlacements(rendered,
                (label, 20D + ((index - 1) / 3) * 160D, 60D + ((index - 1) % 3) * 20D));
        }
        Assert.Equal(2, ColumnPageBandTextPage(rendered, "AFTER").PageNumber);
        AssertColumnBreakPlacements(rendered, ("AFTER", 20D, 20D));
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged)]
    public void HtmlColumns_DeclaredSpanningHeightConstrainsTheFinalPhysicalBandToItsRemainingHeight(
        HtmlRenderIntentProfile profile) {
        HtmlRenderDocument rendered = RenderFiniteSpanningColumns(12, "height:160px", profile);

        Assert.Equal(2, rendered.Pages.Count);
        AssertFiniteSpanningTextOnce(rendered, 12);
        for (int index = 1; index <= 12; index++) {
            string label = FiniteSpanningTailLabel(index);
            int withinPage = (index - 1) % 6;
            Assert.Equal(index <= 6 ? 1 : 2, ColumnPageBandTextPage(rendered, label).PageNumber);
            AssertColumnBreakPlacements(rendered,
                (label, 20D + (withinPage / 3) * 160D,
                    (index <= 6 ? 60D : 20D) + (withinPage % 3) * 20D));
        }
        Assert.Equal(2, ColumnPageBandTextPage(rendered, "AFTER").PageNumber);
        AssertColumnBreakPlacements(rendered, ("AFTER", 20D, 80D));
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged)]
    public void HtmlColumns_SpanningMaximumHeightPreservesOverflowPaintWithoutAdvancingFollowingFlow(
        HtmlRenderIntentProfile profile) {
        HtmlRenderDocument rendered = RenderFiniteSpanningColumns(8, "max-height:100px", profile);

        Assert.Equal(2, rendered.Pages.Count);
        AssertFiniteSpanningTextOnce(rendered, 8);
        for (int index = 1; index <= 6; index++) {
            string label = FiniteSpanningTailLabel(index);
            Assert.Equal(1, ColumnPageBandTextPage(rendered, label).PageNumber);
            AssertColumnBreakPlacements(rendered,
                (label, 20D + ((index - 1) / 3) * 160D, 60D + ((index - 1) % 3) * 20D));
        }
        Assert.Equal(2, ColumnPageBandTextPage(rendered, FiniteSpanningTailLabel(7)).PageNumber);
        Assert.Equal(2, ColumnPageBandTextPage(rendered, FiniteSpanningTailLabel(8)).PageNumber);
        Assert.Equal(2, ColumnPageBandTextPage(rendered, "AFTER").PageNumber);
        AssertColumnBreakPlacements(rendered, (FiniteSpanningTailLabel(7), 20D, 20D),
            (FiniteSpanningTailLabel(8), 20D, 40D), ("AFTER", 20D, 20D));
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData(8, "height:100px", 3, 100D)]
    [InlineData(12, "height:160px", 6, 160D)]
    [InlineData(8, "max-height:100px", 3, 100D)]
    public void HtmlColumns_FiniteSpanningContinuousLayoutRetainsInlineOverflowAndTheAuthoredFlowExtent(
        int tailCount, string sizing, int linesPerColumn, double afterY) {
        HtmlRenderOptions options = ColumnPageBandOptions();
        options.Mode = HtmlRenderMode.Continuous;
        options.ViewportWidth = 800D;
        options.ViewportHeight = 240D;
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(FiniteSpanningSource(tailCount, sizing)), options);

        Assert.Single(rendered.Pages);
        AssertFiniteSpanningTextOnce(rendered, tailCount);
        for (int index = 1; index <= tailCount; index++) {
            AssertColumnBreakPlacements(rendered, (FiniteSpanningTailLabel(index),
                ((index - 1) / linesPerColumn) * 160D, 40D + ((index - 1) % linesPerColumn) * 20D));
        }
        AssertColumnBreakPlacements(rendered, ("AFTER", 0D, afterY));
        AssertColumnBreaksQualified(rendered);
    }

    [Fact]
    public void HtmlColumns_OversizedKeepMakesProgressWithinTheFiniteFinalBand() {
        string html = ColumnBreakSource("<style>@page{size:700px 140px;margin:20px}</style>"
            + "<section style='width:300px;column-count:2;column-gap:20px;column-fill:auto;height:100px'>"
            + "<p>P1</p><p>P2</p><p style='column-span:all'>SPAN</p>"
            + "<div style='break-inside:avoid-column'><p>ATOM1<br>ATOM2<br>ATOM3<br>ATOM4</p></div>"
            + "<p>TAIL</p></section><p>AFTER</p>");
        HtmlRenderOptions options = ColumnPageBandOptions();
        options.MaxColumnCount = 4;

        HtmlRenderDocument rendered = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.DisplayList, options)).Document;

        Assert.Equal(2, rendered.Pages.Count);
        AssertColumnPageBandTextOnce(rendered, new[] { "P1", "P2", "SPAN", "ATOM1", "ATOM2", "ATOM3", "ATOM4", "TAIL", "AFTER" });
        Assert.Equal(2, ColumnPageBandTextPage(rendered, "AFTER").PageNumber);
        AssertColumnBreaksQualified(rendered);
    }

    private static HtmlRenderDocument RenderFiniteSpanningColumns(
        int tailCount, string sizing, HtmlRenderIntentProfile profile) =>
        HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(FiniteSpanningSource(tailCount, sizing)),
            HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, ColumnPageBandOptions())).Document;

    private static string FiniteSpanningSource(int tailCount, string sizing) =>
        ColumnBreakSource("<style>@page{size:700px 140px;margin:20px}</style>"
            + "<section style='width:300px;column-count:2;column-gap:20px;column-fill:auto;" + sizing + "'>"
            + "<p>P1</p><p>P2</p><p style='column-span:all'>SPAN</p>"
            + string.Concat(Enumerable.Range(1, tailCount).Select(index => "<p>" + FiniteSpanningTailLabel(index) + "</p>"))
            + "</section><p>AFTER</p>");

    private static string FiniteSpanningTailLabel(int index) => "TAIL-" + index.ToString("D2");

    private static void AssertFiniteSpanningTextOnce(HtmlRenderDocument rendered, int tailCount) =>
        AssertColumnPageBandTextOnce(rendered, new[] { "P1", "P2", "SPAN", "AFTER" }
            .Concat(Enumerable.Range(1, tailCount).Select(FiniteSpanningTailLabel)));
}

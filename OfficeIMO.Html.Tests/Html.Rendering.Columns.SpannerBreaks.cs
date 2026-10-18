using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, true)]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, false)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, true)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, false)]
    public void HtmlColumns_DirectSpannerForcesItsAuthoredPageBoundary(HtmlRenderIntentProfile profile, bool before) {
        HtmlRenderDocument rendered = RenderSpannerBreakControls(profile,
            "<p>P1</p><p>P2</p><p style='column-span:all;break-" + (before ? "before" : "after")
            + ":page'>SPAN</p><p>T1</p><p>T2</p>", extraStyle: "break-inside:avoid");

        Assert.Equal(2, rendered.Pages.Count);
        AssertColumnPageBandTextOnce(rendered, new[] { "P1", "P2", "SPAN", "T1", "T2", "AFTER" });
        Assert.Equal(before ? 2 : 1, ColumnPageBandTextPage(rendered, "SPAN").PageNumber);
        Assert.Equal(2, ColumnPageBandTextPage(rendered, "T1").PageNumber);
        AssertColumnBreakPlacements(rendered, ("P1", 20D, 20D), ("P2", 180D, 20D),
            ("SPAN", 20D, before ? 20D : 40D), ("T1", 20D, before ? 40D : 20D),
            ("T2", 20D, before ? 60D : 40D), ("AFTER", 20D, before ? 80D : 60D));
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged)]
    public void HtmlColumns_SpannerPageJumpUsesTheActualTerminalPhysicalBand(HtmlRenderIntentProfile profile) {
        string tail = string.Concat(Enumerable.Range(1, 18).Select(index => "<p>T" + index + "</p>"));
        HtmlRenderDocument rendered = RenderSpannerBreakControls(profile,
            "<p>P1</p><p style='break-after:column'>P2</p>"
            + "<p style='column-span:all;break-before:page'>SPAN</p>" + tail,
            "<p>E1</p><p>E2</p>", "column-rule:1px solid blue");

        Assert.Equal(4, rendered.Pages.Count);
        AssertColumnPageBandTextOnce(rendered, new[] { "E1", "E2", "P1", "P2", "SPAN", "AFTER" }
            .Concat(Enumerable.Range(1, 18).Select(index => "T" + index)).ToArray());
        Assert.Equal(2, ColumnPageBandTextPage(rendered, "SPAN").PageNumber);
        Assert.Equal(2, ColumnPageBandTextPage(rendered, "T8").PageNumber);
        Assert.Equal(3, ColumnPageBandTextPage(rendered, "T9").PageNumber);
        Assert.Equal(3, ColumnPageBandTextPage(rendered, "T18").PageNumber);
        Assert.Equal(4, ColumnPageBandTextPage(rendered, "AFTER").PageNumber);
        AssertColumnBreakPlacements(rendered, ("P1", 20D, 60D), ("P2", 180D, 60D),
            ("SPAN", 20D, 20D), ("T1", 20D, 40D), ("T4", 20D, 100D),
            ("T5", 180D, 40D), ("T8", 180D, 100D), ("T9", 20D, 20D),
            ("T14", 180D, 20D), ("T18", 180D, 100D), ("AFTER", 20D, 20D));
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData(true, "avoid")]
    [InlineData(true, "avoid-page")]
    [InlineData(false, "avoid")]
    [InlineData(false, "avoid-page")]
    public void HtmlColumns_SpannerPageKeepProtectsItsAdjacentContent(bool before, string value) {
        HtmlRenderDocument rendered = RenderSpannerBreakControls(HtmlRenderIntentProfile.PrintPaged,
            "<p>P1</p><p>P2</p><p style='column-span:all;break-" + (before ? "before" : "after")
            + ":" + value + "'>SPAN</p><p>T1</p>",
            "<p style='height:" + (before ? "80" : "60") + "px'>ENTRY</p>");

        Assert.Equal(2, rendered.Pages.Count);
        AssertColumnPageBandTextOnce(rendered, new[] { "ENTRY", "P1", "P2", "SPAN", "T1", "AFTER" });
        Assert.Equal(before ? 2 : 1, ColumnPageBandTextPage(rendered, "P1").PageNumber);
        Assert.Equal(2, ColumnPageBandTextPage(rendered, "SPAN").PageNumber);
        Assert.Equal(2, ColumnPageBandTextPage(rendered, "T1").PageNumber);
        AssertColumnBreakPlacements(rendered, ("SPAN", 20D, before ? 40D : 20D),
            ("T1", 20D, before ? 60D : 40D), ("AFTER", 20D, before ? 80D : 60D));
        AssertColumnBreaksQualified(rendered);
    }

    [Fact]
    public void HtmlColumns_AdjacentSpannerPageForcesShareOneBoundaryAndPropagateAtOwnerEdges() {
        HtmlRenderDocument rendered = RenderSpannerBreakControls(HtmlRenderIntentProfile.PrintPaged,
            "<p style='column-span:all;break-before:page;break-after:page'>S1</p>"
            + "<p style='column-span:all;break-before:page;break-after:page'>S2</p>", "<p>ENTRY</p>");

        Assert.Equal(4, rendered.Pages.Count);
        foreach ((string label, int page) in new[] { ("ENTRY", 1), ("S1", 2), ("S2", 3), ("AFTER", 4) })
            Assert.Equal(page, ColumnPageBandTextPage(rendered, label).PageNumber);
        AssertColumnPageBandTextOnce(rendered, new[] { "ENTRY", "S1", "S2", "AFTER" });
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData("break-before:column")]
    [InlineData("break-after:column")]
    [InlineData("break-before:avoid-column")]
    [InlineData("break-after:avoid-column")]
    [InlineData("break-inside:avoid-column")]
    public void HtmlColumns_DirectSpannerColumnControlsHaveNoMatchingOuterColumnContext(string control) {
        HtmlRenderDocument rendered = RenderSpannerBreakControls(HtmlRenderIntentProfile.PrintPaged,
            "<p>P1</p><p>P2</p><p style='column-span:all;" + control + "'>SPAN</p><p>T1</p>");

        Assert.Single(rendered.Pages);
        AssertColumnBreakPlacements(rendered, ("P1", 20D, 20D), ("P2", 180D, 20D),
            ("SPAN", 20D, 40D), ("T1", 20D, 60D), ("AFTER", 20D, 80D));
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData("break-before:page")]
    [InlineData("break-after:page")]
    [InlineData("break-before:avoid-page")]
    [InlineData("break-after:avoid-page")]
    public void HtmlColumns_InactiveSpannerPageControlsRetainContinuousAndSnapshotLayout(string control) {
        string body = "<p>P1</p><p>P2</p><p style='column-span:all;" + control + "'>SPAN</p><p>T1</p>";
        foreach (HtmlRenderIntentProfile profile in new[] {
                     HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderIntentProfile.ScreenSnapshotPaged }) {
            HtmlRenderDocument rendered = RenderSpannerBreakControls(profile, body);
            Assert.Single(rendered.Pages);
            HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
                .OfType<HtmlRenderText>().ToArray();
            Assert.Equal(new[] { "AFTER", "P1", "P2", "SPAN", "T1" },
                text.Select(item => item.Text).OrderBy(value => value, StringComparer.Ordinal));
            AssertColumnBreaksQualified(rendered);
        }
    }

    private static HtmlRenderDocument RenderSpannerBreakControls(HtmlRenderIntentProfile profile,
        string body, string entry = "", string extraStyle = "") {
        string html = ColumnBreakSource("<style>@page{size:700px 140px;margin:20px}</style>" + entry
            + "<section style='width:300px;column-count:2;column-gap:20px;column-fill:auto;" + extraStyle + "'>"
            + body + "</section><p>AFTER</p>");
        return HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
            HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, ColumnPageBandOptions())).Document;
    }
}

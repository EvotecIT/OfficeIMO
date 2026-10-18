using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("column-fill:auto", true, false)]
    [InlineData("column-fill:auto", false, true)]
    [InlineData("column-fill:auto", true, true)]
    [InlineData("column-fill:balance", true, false)]
    [InlineData("column-fill:auto;height:60px", true, false)]
    public void HtmlColumns_ForcedSiblingBoundaryStartsOneNewColumn(
        string columnStyle, bool before, bool after) {
        string body = "<p>BREAK-01</p><p style='" + (after ? "break-after:column" : "")
            + "'>BREAK-02</p><p style='" + (before ? "break-before:column" : "")
            + "'>BREAK-03</p><p>BREAK-04</p><p>BREAK-05</p>";

        HtmlRenderDocument rendered = RenderColumnBreaks(body, columnStyle);

        AssertColumnBreakPlacements(rendered,
            ("BREAK-01", 0D, 0D), ("BREAK-02", 0D, 20D),
            ("BREAK-03", 160D, 0D), ("BREAK-04", 160D, 20D), ("BREAK-05", 160D, 40D),
            ("BREAK-AFTER", 0D, 60D));
        Assert.Equal(60D, FindColumnShape(rendered, "section#break-columns").Height, 3);
        AssertColumnBreaksQualified(rendered);
    }

    [Fact]
    public void HtmlColumns_FirstBeforeAndLastAfterDoNotCreateEmptyColumns() {
        const string body = "<p style='break-before:column'>BREAK-01</p><p>BREAK-02</p>"
            + "<p>BREAK-03</p><p>BREAK-04</p><p style='break-after:column'>BREAK-05</p>";

        HtmlRenderDocument rendered = RenderColumnBreaks(body);

        AssertColumnBreakPlacements(rendered,
            ("BREAK-01", 0D, 0D), ("BREAK-02", 0D, 20D), ("BREAK-03", 0D, 40D),
            ("BREAK-04", 0D, 60D), ("BREAK-05", 0D, 80D), ("BREAK-AFTER", 0D, 100D));
        Assert.Equal(100D, FindColumnShape(rendered, "section#break-columns").Height, 3);
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData(true, 160D)]
    [InlineData(false, 0D)]
    public void HtmlColumns_EmptySiblingHasADistinctForcedBoundary(bool leadingEmpty, double textX) {
        string paragraphs = string.Concat(Enumerable.Range(1, 5).Select(index => "<p>BREAK-0" + index + "</p>"));
        string body = leadingEmpty
            ? "<div id='break-empty' style='break-after:column'></div>" + paragraphs
            : paragraphs + "<div id='break-empty' style='break-before:column'></div>";

        HtmlRenderDocument rendered = RenderColumnBreaks(body);

        AssertColumnBreakPlacements(rendered,
            ("BREAK-01", textX, 0D), ("BREAK-02", textX, 20D), ("BREAK-03", textX, 40D),
            ("BREAK-04", textX, 60D), ("BREAK-05", textX, 80D), ("BREAK-AFTER", 0D, 100D));
        Assert.Equal(100D, FindColumnShape(rendered, "section#break-columns").Height, 3);
        AssertColumnBreaksQualified(rendered);
    }

    [Fact]
    public void HtmlColumns_DistinctEmptyBoundariesAtTheSameHeightAdvanceSeparately() {
        const string body = "<p>BREAK-LEAD</p><div id='break-empty-one' style='break-before:column'></div>"
            + "<div id='break-empty-two' style='break-before:column'></div><p>BREAK-TAIL</p>";

        HtmlRenderDocument rendered = RenderColumnBreaks(body);

        AssertColumnBreakPlacements(rendered,
            ("BREAK-LEAD", 0D, 0D), ("BREAK-TAIL", 320D, 0D), ("BREAK-AFTER", 0D, 20D));
        Assert.Equal(20D, FindColumnShape(rendered, "section#break-columns").Height, 3);
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData("break-before", "avoid-column", 2)]
    [InlineData("break-after", "avoid-column", 2)]
    [InlineData("break-before", "avoid-page", 3)]
    [InlineData("break-after", "avoid-page", 3)]
    [InlineData("break-before", "avoid", 2)]
    [InlineData("break-after", "avoid", 2)]
    public void HtmlColumns_BoundaryAvoidanceAppliesOnlyToItsFragmentationScope(
        string property, string value, int leftCount) {
        int styledIndex = property == "break-before" ? 4 : 3;
        string body = string.Concat(Enumerable.Range(1, 5).Select(index => "<p style='"
            + (index == styledIndex ? property + ":" + value : "") + "'>BREAK-0" + index + "</p>"));

        HtmlRenderDocument rendered = RenderColumnBreaks(body, "column-fill:auto;height:60px");

        for (int index = 1; index <= 5; index++) {
            int inColumn = index <= leftCount ? index - 1 : index - leftCount - 1;
            AssertColumnBreakPlacements(rendered,
                ("BREAK-0" + index, index <= leftCount ? 0D : 160D, 20D * inColumn));
        }
        AssertColumnBreakPlacements(rendered, ("BREAK-AFTER", 0D, 60D));
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData("avoid-column", 160D, 0D)]
    [InlineData("avoid-page", 0D, 40D)]
    [InlineData("avoid", 160D, 0D)]
    public void HtmlColumns_InsideAvoidanceKeepsAFittingGroupInItsMatchingScope(
        string value, double groupX, double groupY) {
        string body = "<p>BREAK-01</p><p>BREAK-02</p><div style='break-inside:" + value
            + "'><p>BREAK-03</p><p>BREAK-04</p></div><p>BREAK-05</p>";

        HtmlRenderDocument rendered = RenderColumnBreaks(body, "column-fill:auto;height:60px");

        AssertColumnBreakPlacements(rendered,
            ("BREAK-01", 0D, 0D), ("BREAK-02", 0D, 20D), ("BREAK-03", groupX, groupY),
            ("BREAK-04", 160D, value == "avoid-page" ? 0D : 20D),
            ("BREAK-05", 160D, value == "avoid-page" ? 20D : 40D), ("BREAK-AFTER", 0D, 60D));
        AssertColumnBreaksQualified(rendered);
    }

    [Fact]
    public void HtmlColumns_OversizedColumnKeepStartsFreshThenAllowsProgress() {
        const string body = "<p>BREAK-01</p><div style='break-inside:avoid-column'>"
            + "<p>BREAK-02</p><p>BREAK-03</p><p>BREAK-04</p><p>BREAK-05</p><p>BREAK-06</p></div>";

        HtmlRenderDocument rendered = RenderColumnBreaks(body, "column-fill:auto;height:60px");

        AssertColumnBreakPlacements(rendered,
            ("BREAK-01", 0D, 0D), ("BREAK-02", 160D, 0D), ("BREAK-03", 160D, 20D),
            ("BREAK-04", 160D, 40D), ("BREAK-05", 320D, 0D), ("BREAK-06", 320D, 20D),
            ("BREAK-AFTER", 0D, 60D));
        AssertColumnBreaksQualified(rendered);
    }

    [Fact]
    public void HtmlColumns_AvoidAfterKeepsTheFinalParagraphLineWithFollowingContent() {
        const string body = "<p>BREAK-LEAD</p><p style='break-after:avoid-column'>"
            + "BREAK-A01<br>BREAK-A02<br>BREAK-A03</p><p>BREAK-B01<br>BREAK-B02<br>BREAK-B03</p>";

        HtmlRenderDocument rendered = RenderColumnBreaks(body, "column-fill:auto;height:80px");

        AssertColumnBreakPlacements(rendered,
            ("BREAK-LEAD", 0D, 0D), ("BREAK-A01", 0D, 20D), ("BREAK-A02", 0D, 40D),
            ("BREAK-A03", 160D, 0D), ("BREAK-B01", 160D, 20D), ("BREAK-B02", 160D, 40D),
            ("BREAK-B03", 160D, 60D), ("BREAK-AFTER", 0D, 80D));
        AssertColumnBreaksQualified(rendered);
    }

    [Fact]
    public void HtmlColumns_AvoidAfterDoesNotBacktrackWhenFollowingContentFits() {
        const string body = "<p>BREAK-LEAD</p><p style='break-after:avoid-column'>"
            + "BREAK-A01<br>BREAK-A02<br>BREAK-A03</p><p>BREAK-B01<br>BREAK-B02<br>BREAK-B03</p>";

        HtmlRenderDocument rendered = RenderColumnBreaks(body, "column-fill:auto;height:120px");

        AssertColumnBreakPlacements(rendered,
            ("BREAK-LEAD", 0D, 0D), ("BREAK-A01", 0D, 20D), ("BREAK-A02", 0D, 40D),
            ("BREAK-A03", 0D, 60D), ("BREAK-B01", 0D, 80D), ("BREAK-B02", 0D, 100D),
            ("BREAK-B03", 160D, 0D), ("BREAK-AFTER", 0D, 120D));
        AssertColumnBreaksQualified(rendered);
    }

    [Fact]
    public void HtmlColumns_OverlappingBoundaryKeepsMoveTheCompleteChainTogether() {
        const string body = "<p>BREAK-LEAD01</p><p>BREAK-LEAD02</p>"
            + "<p style='break-after:avoid-column'>BREAK-A</p>"
            + "<p style='break-after:avoid-column'>BREAK-B</p><p>BREAK-C</p>";

        HtmlRenderDocument rendered = RenderColumnBreaks(body, "column-fill:auto;height:80px");

        AssertColumnBreakPlacements(rendered,
            ("BREAK-LEAD01", 0D, 0D), ("BREAK-LEAD02", 0D, 20D),
            ("BREAK-A", 160D, 0D), ("BREAK-B", 160D, 20D), ("BREAK-C", 160D, 40D),
            ("BREAK-AFTER", 0D, 80D));
        AssertColumnBreaksQualified(rendered);
    }

    [Fact]
    public void HtmlColumns_ForcedBoundaryOverridesAnAncestorColumnKeep() {
        const string body = "<p>BREAK-LEAD</p><div style='break-inside:avoid-column'><p>BREAK-ONE</p>"
            + "<p style='break-before:column'>BREAK-TWO</p><p>BREAK-THREE</p></div>";

        HtmlRenderDocument rendered = RenderColumnBreaks(body, "column-fill:auto;height:80px");

        AssertColumnBreakPlacements(rendered,
            ("BREAK-LEAD", 0D, 0D), ("BREAK-ONE", 0D, 20D), ("BREAK-TWO", 160D, 0D),
            ("BREAK-THREE", 160D, 20D), ("BREAK-AFTER", 0D, 80D));
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData("break-before:column")]
    [InlineData("break-after:column")]
    [InlineData("break-before:avoid-column")]
    [InlineData("break-after:avoid-column")]
    [InlineData("break-inside:avoid-column")]
    public void HtmlPagedMedia_ColumnBreakValuesOutsideColumnsDoNotCreatePages(string declaration) {
        HtmlRenderOptions options = ColumnPageBandOptions();
        options.Mode = HtmlRenderMode.Paged;
        options.HonorCssPageRules = false;
        options.PageSize = new OfficePageSize(400D / HtmlRenderOptions.CssPixelsPerInch,
            200D / HtmlRenderOptions.CssPixelsPerInch);
        string html = ColumnBreakSource("<p>BREAK-ONE</p><p style='" + declaration
            + "'>BREAK-TWO</p><p>BREAK-THREE</p>");

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        Assert.Single(rendered.Pages);
        AssertColumnBreakPlacements(rendered,
            ("BREAK-ONE", 0D, 0D), ("BREAK-TWO", 0D, 20D), ("BREAK-THREE", 0D, 40D));
        AssertColumnBreaksQualified(rendered);
    }

    private static HtmlRenderDocument RenderColumnBreaks(string body, string columnStyle = "column-fill:auto") {
        HtmlRenderOptions options = ColumnPageBandOptions();
        options.Mode = HtmlRenderMode.Continuous;
        options.ViewportWidth = 800D;
        options.ViewportHeight = 240D;
        options.HonorCssPageRules = false;
        return HtmlRenderTestDriver.Render(ColumnBreakSource("<section id='break-columns' style='width:300px;"
            + "column-count:2;column-gap:20px;background:#eee;" + columnStyle + "'>" + body
            + "</section><p>BREAK-AFTER</p>"), options);
    }

    private static string ColumnBreakSource(string body) =>
        "<!doctype html><style>html,body{margin:0;font:14px/20px PageBandWitness}"
        + "p{margin:0;orphans:1;widows:1}</style>" + body;

    private static void AssertColumnBreakPlacements(HtmlRenderDocument rendered,
        params (string Label, double X, double Y)[] expected) {
        HtmlRenderText[] text = rendered.Pages.SelectMany(ColumnPageBandText).ToArray();
        foreach ((string label, double x, double y) in expected) {
            HtmlRenderText actual = Assert.Single(text, item => item.Text == label);
            Assert.Equal(x, actual.X, 3);
            Assert.Equal(y, actual.LayoutY, 3);
        }
    }

    private static void AssertColumnBreaksQualified(HtmlRenderDocument rendered) =>
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.MultiColumnValueUnsupported
            || diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment
            || diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
}

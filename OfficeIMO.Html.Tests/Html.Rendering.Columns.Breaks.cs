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

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlPagedMedia_InapplicableColumnForceDoesNotMaskTableRowPageBreak(bool tableEntry) {
        HtmlRenderOptions options = ColumnPageBandOptions();
        options.Mode = HtmlRenderMode.Paged;
        options.HonorCssPageRules = false;
        options.PageSize = new OfficePageSize(400D / HtmlRenderOptions.CssPixelsPerInch,
            200D / HtmlRenderOptions.CssPixelsPerInch);
        string table = tableEntry
            ? "<table style='break-before:column'><tr style='break-before:page'><td>BREAK-ONE</td></tr><tr>"
            : "<table><tr style='break-after:column'><td>BREAK-ONE</td></tr><tr style='break-before:page'>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(ColumnBreakSource(
            "<p>BREAK-LEAD</p>" + table + "<td>BREAK-TWO</td></tr></table>"), options);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Single(ColumnPageBandText(rendered.Pages[0]), item => item.Text == "BREAK-LEAD");
        Assert.Single(ColumnPageBandText(rendered.Pages[tableEntry ? 1 : 0]), item => item.Text == "BREAK-ONE");
        Assert.Single(ColumnPageBandText(rendered.Pages[1]), item => item.Text == "BREAK-TWO");
        AssertColumnBreaksQualified(rendered);
    }

    [Theory]
    [InlineData("<div style='display:flex'><div style='break-before:column'>BREAK-TWO</div></div>")]
    [InlineData("<div style='display:grid'><div style='break-before:column'>BREAK-TWO</div></div>")]
    [InlineData("<div style='column-count:2;font-size:8px'><p style='break-before:column'>BREAK-TWO</p></div>")]
    [InlineData("<p style='break-before:column'>BREAK-TWO</p><div style='column-span:all'>BREAK-SPAN</div>")]
    [InlineData("<table><tr style='break-before:column'><td>BREAK-TWO</td></tr></table>")]
    public void HtmlColumns_UnqualifiedBreakContextsReportLoss(string body) {
        HtmlRenderDocument rendered = RenderColumnBreaks("<p>BREAK-ONE</p>" + body);

        Assert.Single(rendered.Pages.SelectMany(ColumnPageBandText), item => item.Text == "BREAK-TWO");
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.MultiColumnValueUnsupported);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

    [Fact]
    public void HtmlColumns_ParagraphMarginCutPreservesFittingLinesAndDiscardsOnlyUnpaintedTail() {
        const string body = "<p style='height:40px'>BREAK-LEAD</p>"
            + "<p style='margin-bottom:20px;orphans:2;widows:2'>BREAK-ONE<br>BREAK-TWO<br>BREAK-THREE</p>"
            + "<p>BREAK-FOUR</p>";

        HtmlRenderDocument rendered = RenderColumnBreaks(body, "column-fill:auto;height:100px");

        AssertColumnBreakPlacements(rendered, ("BREAK-LEAD", 0D, 0D), ("BREAK-ONE", 0D, 40D),
            ("BREAK-TWO", 0D, 60D), ("BREAK-THREE", 0D, 80D), ("BREAK-FOUR", 160D, 0D),
            ("BREAK-AFTER", 0D, 100D));
        AssertColumnBreaksQualified(rendered);
    }

    [Fact]
    public void HtmlColumns_InterruptedInlineBoundaryCoalescesBeforeAndAfterForces() {
        const string body = "<span><span style='display:block;break-after:column'>BREAK-ONE</span>"
            + "<span style='display:block;break-before:column'>BREAK-TWO</span></span>";

        HtmlRenderDocument rendered = RenderColumnBreaks(body);

        AssertColumnBreakPlacements(rendered, ("BREAK-ONE", 0D, 0D), ("BREAK-TWO", 160D, 0D),
            ("BREAK-AFTER", 0D, 20D));
        AssertColumnBreaksQualified(rendered);
    }

    [Fact]
    public void HtmlColumnNotes_PageOnlyNestedKeepDoesNotEnlargeColumnCallReservation() {
        HtmlRenderDocument Render(string inside) {
            HtmlRenderOptions options = ColumnPageBandOptions();
            options.Mode = HtmlRenderMode.Paged;
            options.HonorCssPageRules = true;
            return HtmlRenderTestDriver.Render(ColumnBreakSource(
                "<style>@page{size:700px 140px;margin:20px}.note{float:footnote;float-reference:column;font-size:10px;line-height:12px}</style>"
                + "<section style='width:300px;column-count:2;column-gap:20px;column-fill:auto;height:100px'>"
                + "<div><div style='break-inside:" + inside + "'><p>CALL<span class='note'>NOTEONE<br>NOTETWO<br>NOTETHREE</span></p>"
                + "<p>BODYTWO</p><p>BODYTHREE</p></div></div></section><p>AFTER</p>"), options);
        }
        static (int Page, string Text, double X, double Y)[] Placements(HtmlRenderDocument document) =>
            document.Pages.SelectMany(page => ColumnPageBandText(page)
                .Select(text => (page.PageNumber, text.Text, text.X, text.LayoutY))).ToArray();

        HtmlRenderDocument control = Render("auto");
        HtmlRenderDocument pageKeep = Render("avoid-page");

        Assert.Equal(Placements(control), Placements(pageKeep));
        HtmlRenderText[] text = pageKeep.Pages.SelectMany(ColumnPageBandText).ToArray();
        Assert.Single(text, item => item.Text == "NOTEONE");
        Assert.Single(text, item => item.Text == "NOTETWO");
        Assert.Single(text, item => item.Text == "NOTETHREE");
        Assert.DoesNotContain(text, item => item.Text.Contains("cont.", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlColumns_DiscardedMarginRetainsOneAdvanceBeforeFollowingSmallBlock(bool forced) {
        string body = "<p style='height:30px'>BREAK-LEAD</p><p style='margin-bottom:20px;orphans:2;widows:2;"
            + (forced ? "break-after:column" : "") + "'>BREAK-ONE<br>BREAK-TWO<br>BREAK-THREE</p>"
            + "<div id='break-spacer' style='height:5px;background:#174b72;font-size:0'></div><p>BREAK-FOUR</p>";

        HtmlRenderDocument rendered = RenderColumnBreaks(body, "column-fill:auto;height:100px");

        AssertColumnBreakPlacements(rendered, ("BREAK-LEAD", 0D, 0D), ("BREAK-ONE", 0D, 30D),
            ("BREAK-TWO", 0D, 50D), ("BREAK-THREE", 0D, 70D), ("BREAK-FOUR", 160D, 5D),
            ("BREAK-AFTER", 0D, 100D));
        HtmlRenderShape spacer = FindColumnShape(rendered, "div#break-spacer");
        Assert.Equal(160D, spacer.X, 3);
        Assert.Equal(0D, spacer.Y, 3);
        Assert.Equal(5D, spacer.Height, 3);
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

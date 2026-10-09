using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, 50D, 150D)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, 80D, 120D)]
    [InlineData(HtmlRenderIntentProfile.ScreenSnapshotPaged, 80D, 120D)]
    public void HtmlTablePercentageHeights_NamedProfilesUseTheirEffectiveMediaAndRetainKnownGeometry(HtmlRenderIntentProfile profile, double summaryHeight, double detailHeight) {
        string html = TableGeometrySource("<style>@media print{#summary{height:25%}#detail{height:75%}}"
            + "@media screen{#summary{height:40%}#detail{height:60%}}</style>"
            + "<table style='height:200px;width:240px;margin:0;border-spacing:0'><tr><td id='summary' style='padding:0;background:lime'>Summary</td></tr>"
            + "<tr><td id='detail' style='padding:0;background:lime'>Detail</td></tr></table><div id='next' style='height:20px;background:red'>Next</div>");
        var options = TableGeometryOptions();
        options.PageSize = new OfficePageSize(600D / 96D, 300D / 96D);
        options.HonorCssPageRules = false;
        HtmlRenderDocument rendered = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
            HtmlRenderRequest.Create(profile, options: options)).Document;

        Assert.Single(rendered.Pages);
        HtmlRenderShape Shape(string source) => Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderShape>(),
            shape => shape.Source == source && shape.Shape.FillColor.HasValue);
        Assert.Equal(summaryHeight, Shape("td#summary").Height, 3);
        Assert.Equal(detailHeight, Shape("td#detail").Height, 3);
        Assert.Equal(200D, Shape("div#next").Y, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("tr", false, false)]
    [InlineData("td", false, false)]
    [InlineData("tr", false, true)]
    [InlineData("td", false, true)]
    [InlineData("tr", true, true)]
    [InlineData("td", true, true)]
    public void HtmlTablePercentageHeights_DefiniteGridAllocatesRowAndCellShares(string owner, bool cssTable, bool explicitGroup) {
        string TableTag(string native, string display) => cssTable ? "div style='display:" + display + "'" : native;
        string Cell(string id, string height, string text) => cssTable
            ? "<div id='" + id + "' style='display:table-cell;padding:0;background:lime;" + (owner == "td" ? "height:" + height : "") + "'>" + text + "</div>"
            : "<td id='" + id + "' style='padding:0;background:lime;" + (owner == "td" ? "height:" + height : "") + "'>" + text + "</td>";
        string Row(string height, string cell) => cssTable
            ? "<div style='display:table-row;" + (owner == "tr" ? "height:" + height : "") + "'>" + cell + "</div>"
            : "<tr style='" + (owner == "tr" ? "height:" + height : "") + "'>" + cell + "</tr>";
        string rows = Row("25%", Cell("summary", "25%", "Summary")) + Row("75%", Cell("detail", "75%", "Detail"));
        if (explicitGroup) rows = "<" + TableTag("tbody", "table-row-group") + ">" + rows + (cssTable ? "</div>" : "</tbody>");
        string html = TableGeometrySource((cssTable ? "<div style='display:table;" : "<table style='")
            + "height:200px;width:240px;margin:0;border-spacing:0'>" + rows + (cssTable ? "</div>" : "</table>")
            + "<div id='next' style='height:20px;background:red'>Next</div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Equal(50D, TableGeometryShape(rendered, (cssTable ? "div" : "td") + "#summary").Height, 3);
        Assert.Equal(150D, TableGeometryShape(rendered, (cssTable ? "div" : "td") + "#detail").Height, 3);
        Assert.Equal(200D, TableGeometryShape(rendered, "div#next").Y, 3);
        Assert.False(rendered.HasLoss);
    }

    [Fact]
    public void HtmlTablePercentageHeights_RowAndCellRequestsUseTheLargerShare() {
        string html = TableGeometrySource("<table style='height:200px;width:240px;margin:0;border-spacing:0'>"
            + "<tr style='height:25%'><td id='summary' style='height:75%;padding:0;background:lime'>Summary</td></tr>"
            + "<tr style='height:25%'><td id='detail' style='height:10%;padding:0;background:lime'>Detail</td></tr></table>"
            + "<div id='next' style='height:20px;background:red'>Next</div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Equal(150D, TableGeometryShape(rendered, "td#summary").Height, 3);
        Assert.Equal(50D, TableGeometryShape(rendered, "td#detail").Height, 3);
        Assert.Equal(200D, TableGeometryShape(rendered, "div#next").Y, 3);
        Assert.False(rendered.HasLoss);
    }

    [Fact]
    public void HtmlTablePercentageHeights_ContentMinimumRedistributesSharesWithoutClippingOrGrowingTheTable() {
        string html = TableGeometrySource("<table style='height:200px;width:240px;margin:0;border-spacing:0'>"
            + "<tr><td id='summary' style='height:25%;padding:0;background:lime'><div id='content' style='height:80px;background:yellow'>Summary</div></td></tr>"
            + "<tr><td id='detail' style='height:75%;padding:0;background:lime'>Detail</td></tr></table>"
            + "<div id='next' style='height:20px;background:red'>Next</div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Equal(80D, TableGeometryShape(rendered, "td#summary").Height, 3);
        Assert.Equal(80D, TableGeometryShape(rendered, "div#content").Height, 3);
        Assert.Equal(120D, TableGeometryShape(rendered, "td#detail").Height, 3);
        Assert.Equal(200D, TableGeometryShape(rendered, "div#next").Y, 3);
        Assert.False(rendered.HasLoss);
    }

    [Theory]
    [InlineData("", 50D, 150D)]
    [InlineData("height:50%", 200D / 3D, 400D / 3D)]
    [InlineData("height:150px", 50D, 150D)]
    public void HtmlTablePercentageHeights_AllocateRemainingHeightToAutoRowsOrScalePercentageShares(string detailStyle, double summaryHeight, double detailHeight) {
        string html = TableGeometrySource("<table style='height:200px;width:240px;margin:0;border-spacing:0'>"
            + "<tr><td id='summary' style='height:25%;padding:0;background:lime'>Summary</td></tr>"
            + "<tr><td id='detail' style='padding:0;background:lime;" + detailStyle + "'>Detail</td></tr></table>"
            + "<div id='next' style='height:20px;background:red'>Next</div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Equal(summaryHeight, TableGeometryShape(rendered, "td#summary").Height, 3);
        Assert.Equal(detailHeight, TableGeometryShape(rendered, "td#detail").Height, 3);
        Assert.Equal(200D, TableGeometryShape(rendered, "div#next").Y, 3);
        Assert.False(rendered.HasLoss);
    }

    [Theory]
    [InlineData("border-box", "border-box", 240D, 52D, 156D)]
    [InlineData("border-box", "content-box", 240D, 52D, 156D)]
    [InlineData("content-box", "border-box", 260D, 57D, 171D)]
    [InlineData("content-box", "content-box", 260D, 57D, 171D)]
    public void HtmlTablePercentageHeights_TableBoxSizingDeductsInsetsAndSpacingFromTheRowGrid(string tableSizing, string cellSizing, double tableHeight, double summaryHeight, double detailHeight) {
        string cellStyle = "padding:4px;border:1px solid navy;background:lime;box-sizing:" + cellSizing + ";";
        string html = TableGeometrySource("<table id='report' style='height:240px;width:240px;margin:0;background:white;border:5px solid black;padding:5px;border-spacing:4px;box-sizing:"
            + tableSizing + "'><tr><td id='summary' style='" + cellStyle + "height:25%'>Summary</td></tr>"
            + "<tr><td id='detail' style='" + cellStyle + "height:75%'>Detail</td></tr></table>"
            + "<div id='next' style='height:20px;background:red'>Next</div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Equal(tableHeight, TableGeometryShape(rendered, "table#report").Height, 3);
        Assert.Equal(summaryHeight, TableGeometryShape(rendered, "td#summary").Height, 3);
        Assert.Equal(detailHeight, TableGeometryShape(rendered, "td#detail").Height, 3);
        Assert.Equal(14D, TableGeometryShape(rendered, "td#summary").Y, 3);
        Assert.Equal(tableHeight, TableGeometryShape(rendered, "div#next").Y, 3);
        Assert.False(rendered.HasLoss);
    }

    [Fact]
    public void HtmlTablePercentageHeights_MixedNonzeroMathUsesAutoAfterVariableResolution() {
        string html = TableGeometrySource("<table style='--summary:25%;height:200px;width:240px;margin:0;border-spacing:0'>"
            + "<tr><td id='summary' style='height:75%;height:calc(var(--summary) + 10px);padding:0;background:lime'>Summary</td></tr>"
            + "<tr><td id='detail' style='height:calc(75% - 10px);padding:0;background:lime'>Detail</td></tr></table>"
            + "<div id='next' style='height:20px;background:red'>Next</div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Equal(100D, TableGeometryShape(rendered, "td#summary").Height, 3);
        Assert.Equal(100D, TableGeometryShape(rendered, "td#detail").Height, 3);
        Assert.Equal(200D, TableGeometryShape(rendered, "div#next").Y, 3);
        Assert.False(rendered.HasLoss);
    }

    [Theory]
    [InlineData("tr", "0%")]
    [InlineData("td", "0%")]
    [InlineData("tr", "0px")]
    [InlineData("td", "0px")]
    [InlineData("tr", "10px - 10px")]
    [InlineData("td", "10px - 10px")]
    public void HtmlTablePercentageHeights_PureOrZeroLengthCalculatedSharesKeepTheirPercentages(string owner, string absoluteTerm) {
        string first = "height:calc(25% + " + absoluteTerm + ")";
        string second = "height:calc(75% + " + absoluteTerm + ")";
        string html = TableGeometrySource("<table style='height:200px;width:240px;margin:0;border-spacing:0'>"
            + "<tr style='" + (owner == "tr" ? first : "") + "'><td id='summary' style='padding:0;background:lime;" + (owner == "td" ? first : "") + "'>Summary</td></tr>"
            + "<tr style='" + (owner == "tr" ? second : "") + "'><td id='detail' style='padding:0;background:lime;" + (owner == "td" ? second : "") + "'>Detail</td></tr></table>"
            + "<div id='next' style='height:20px;background:red'>Next</div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Equal(50D, TableGeometryShape(rendered, "td#summary").Height, 3);
        Assert.Equal(150D, TableGeometryShape(rendered, "td#detail").Height, 3);
        Assert.Equal(200D, TableGeometryShape(rendered, "div#next").Y, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("min(25%, 10px)", "max(75%, 10px)")]
    [InlineData("calc(0% + 80px)", "calc(0% + 120px)")]
    public void HtmlTablePercentageHeights_MixedComparisonAndZeroPercentageMathKeepAutoLossFree(string first, string second) {
        string html = TableGeometrySource("<table style='height:200px;width:240px;margin:0;border-spacing:0'>"
            + "<tr><td id='summary' style='height:" + first + ";padding:0;background:lime'>Summary</td></tr>"
            + "<tr><td id='detail' style='height:" + second + ";padding:0;background:lime'>Detail</td></tr></table>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Equal(100D, TableGeometryShape(rendered, "td#summary").Height, 3);
        Assert.Equal(100D, TableGeometryShape(rendered, "td#detail").Height, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlTablePercentageHeights_FixedTableCellAndColumnMixedWidthsUseTheSameAutoRule(bool column) {
        string html = TableGeometrySource("<table style='width:400px;table-layout:fixed;margin:0;border-spacing:0'>"
            + (column ? "<colgroup><col style='width:calc(25% + 20px)'><col></colgroup>" : "")
            + "<tr><td id='summary' style='padding:0;background:lime;" + (column ? "" : "width:calc(25% + 20px)") + "'>Summary</td>"
            + "<td id='detail' style='padding:0;background:lime'>Detail</td></tr></table>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Equal(200D, TableGeometryShape(rendered, "td#summary").Width, 3);
        Assert.Equal(200D, TableGeometryShape(rendered, "td#detail").Width, 3);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlTablePercentageHeights_TableRootsAndOrdinaryBlocksRetainValidMixedMath() {
        string html = TableGeometrySource("<div id='ordinary' style='width:calc(50% + 10px);background:red'>Ordinary</div>"
            + "<table id='report' style='width:calc(25% + 20px);table-layout:fixed;margin:0;border-spacing:0;background:white'>"
            + "<tr><td id='summary' style='padding:0;background:lime'>Summary</td><td id='detail' style='padding:0;background:lime'>Detail</td></tr></table>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Equal(310D, TableGeometryShape(rendered, "div#ordinary").Width, 3);
        Assert.Equal(170D, TableGeometryShape(rendered, "table#report").Width, 3);
        Assert.Equal(85D, TableGeometryShape(rendered, "td#summary").Width, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("height:25%;height:invalid", 100D, 100D)]
    [InlineData("height:25%;height:initial", 150D, 50D)]
    public void HtmlTablePercentageHeights_InvalidDeclarationsAndCssWideResetsKeepTheEffectiveHeight(string summaryStyle, double summaryHeight, double detailHeight) {
        string html = TableGeometrySource("<table style='height:200px;width:240px;margin:0;border-spacing:0'><tr>"
            + "<td id='summary' style='padding:0;background:lime;" + summaryStyle + "'>Summary</td></tr>"
            + "<tr><td id='detail' style='height:25%;padding:0;background:lime'>Detail</td></tr></table>"
            + "<div id='next' style='height:20px;background:red'>Next</div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Equal(summaryHeight, TableGeometryShape(rendered, "td#summary").Height, 3);
        Assert.Equal(detailHeight, TableGeometryShape(rendered, "td#detail").Height, 3);
        Assert.Equal(200D, TableGeometryShape(rendered, "div#next").Y, 3);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlTablePercentageHeights_IndefiniteTableKeepsLegitimateAutoSizingWithoutLoss() {
        string html = TableGeometrySource("<table style='width:240px;margin:0;border-spacing:0'><tbody><tr style='height:25%'>"
            + "<td id='summary' style='height:25%;padding:0;background:lime'>Summary</td></tr><tr style='height:75%'>"
            + "<td id='detail' style='height:75%;padding:0;background:lime'>Detail</td></tr></tbody></table>"
            + "<div id='next' style='height:20px;background:red'>Next</div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Equal(20D, TableGeometryShape(rendered, "td#summary").Height, 3);
        Assert.Equal(20D, TableGeometryShape(rendered, "td#detail").Height, 3);
        Assert.Equal(40D, TableGeometryShape(rendered, "div#next").Y, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(true, 20, 2, 200D)]
    [InlineData(false, 20, 2, 200D)]
    [InlineData(true, 60, 3, 0D)]
    [InlineData(false, 60, 3, 0D)]
    public void HtmlTablePercentageHeights_ContinuationKeepsSharesRepeatedGroupsAndFollowingGeometry(bool percentages, int followingHeight, int pages, double followingY) {
        string html = TableGeometrySource("<table style='height:400px;width:180px;margin:0;border-spacing:0'>"
            + "<thead><tr style='height:5%'><th style='padding:0'>Header</th></tr></thead><tbody>"
            + "<tr style='height:20%'><td id='body1' style='padding:0;background:lime'>Body1</td></tr>"
            + "<tr style='height:30%'><td id='body2' style='padding:0;background:lime'>Body2</td></tr>"
            + "<tr style='height:40%'><td id='body3' style='padding:0;background:lime'>Body3</td></tr></tbody>"
            + "<tfoot><tr style='height:5%'><td style='padding:0'>Footer</td></tr></tfoot></table>"
            + "<div id='next' style='height:" + followingHeight + "px;background:red;break-inside:avoid'>Next</div>");
        if (!percentages) html = html.Replace("height:5%", "height:20px").Replace("height:20%", "height:80px")
            .Replace("height:30%", "height:120px").Replace("height:40%", "height:160px");
        var options = TableGeometryOptions();
        options.Mode = HtmlRenderMode.Paged;
        options.PageSize = new OfficePageSize(200D / 96D, 250D / 96D);
        options.HonorCssPageRules = false;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        Assert.Equal(pages, rendered.Pages.Count);
        Assert.Equal(80D, TableGeometryShape(rendered, "td#body1").Height, 3);
        Assert.Equal(120D, TableGeometryShape(rendered, "td#body2").Height, 3);
        Assert.Equal(160D, TableGeometryShape(rendered, "td#body3").Height, 3);
        Assert.Equal(followingY, TableGeometryShape(rendered, "div#next").Y, 3);
        Assert.All(rendered.Pages.Take(2), page => Assert.Single(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "Header"));
        Assert.All(rendered.Pages.Take(2), page => Assert.Single(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "Footer"));
        Assert.Equal(2, rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().Count(text => text.Text == "Header"));
        Assert.Equal(2, rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().Count(text => text.Text == "Footer"));
        foreach (string text in new[] { "Body1", "Body2", "Body3", "Next" }) {
            Assert.Single(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), item => item.Text == text);
        }
        Assert.False(rendered.HasLoss);
    }

    [Fact]
    public void HtmlTablePercentageHeights_FinalRepeatedFooterLeavesRoomForFollowingContentWithoutAHeader() {
        string html = TableGeometrySource("<table style='height:380px;width:180px;margin:0;border-spacing:0'><tbody>"
            + "<tr style='height:80px'><td style='padding:0'>Body1</td></tr>"
            + "<tr style='height:120px'><td style='padding:0'>Body2</td></tr>"
            + "<tr style='height:160px'><td style='padding:0'>Body3</td></tr></tbody>"
            + "<tfoot><tr style='height:20px'><td style='padding:0'>Footer</td></tr></tfoot></table>"
            + "<div id='next' style='height:20px;background:red'>Next</div>");
        var options = TableGeometryOptions();
        options.Mode = HtmlRenderMode.Paged;
        options.PageSize = new OfficePageSize(200D / 96D, 250D / 96D);
        options.HonorCssPageRules = false;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Equal(180D, TableGeometryShape(rendered, "div#next").Y, 3);
        Assert.All(rendered.Pages, page => Assert.Single(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "Footer"));
        foreach (string text in new[] { "Body1", "Body2", "Body3", "Next" }) {
            Assert.Single(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), item => item.Text == text);
        }
        Assert.False(rendered.HasLoss);
    }

    [Theory]
    [InlineData("break-before:page", 200D)]
    [InlineData("page:next", 300D)]
    public void HtmlTablePercentageHeights_FinalRepeatedFooterHonorsFollowingPageBoundaries(string followingStyle, double followingPageWidth) {
        string html = TableGeometrySource("<style>@page{size:200px 250px;margin:0}@page next{size:300px 250px;margin:0}</style>"
            + "<table style='height:400px;width:180px;margin:0;border-spacing:0'>"
            + "<thead><tr style='height:20px'><th style='padding:0'>Header</th></tr></thead><tbody>"
            + "<tr style='height:80px'><td style='padding:0'>Body1</td></tr>"
            + "<tr style='height:120px'><td style='padding:0'>Body2</td></tr>"
            + "<tr style='height:160px'><td style='padding:0'>Body3</td></tr></tbody>"
            + "<tfoot><tr style='height:20px'><td style='padding:0'>Footer</td></tr></tfoot></table>"
            + "<div id='next' style='height:20px;background:red;" + followingStyle + "'>Next</div>");
        var options = TableGeometryOptions();
        options.Mode = HtmlRenderMode.Paged;
        options.HonorCssPageRules = true;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        Assert.Equal(3, rendered.Pages.Count);
        Assert.Equal(followingPageWidth, rendered.Pages[2].Width, 3);
        Assert.Equal(0D, TableGeometryShape(rendered, "div#next").Y, 3);
        Assert.Single(rendered.Pages[2].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Next");
        Assert.Equal(2, rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().Count(text => text.Text == "Header"));
        Assert.Equal(2, rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().Count(text => text.Text == "Footer"));
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlTablePercentageHeights_FollowingInlineContentAndRunningStringStayOnTheFinalFooterPage() {
        string html = TableGeometrySource("<style>@page{size:200px 270px;margin:0 0 20px;@bottom-center{content:string(chapter,last)}}</style>"
            + "<table style='height:400px;width:180px;margin:0;border-spacing:0'>"
            + "<thead><tr style='height:20px'><th style='padding:0'>Header</th></tr></thead><tbody>"
            + "<tr style='height:80px'><td style='padding:0'>Body1</td></tr>"
            + "<tr style='height:120px'><td style='padding:0'>Body2</td></tr>"
            + "<tr style='height:160px'><td style='padding:0'>Body3</td></tr></tbody>"
            + "<tfoot><tr style='height:20px'><td style='padding:0'>Footer</td></tr></tfoot></table>"
            + "<div id='next' style='string-set:chapter content();margin:0'>Next <span>inline</span> continuation</div>");
        var options = TableGeometryOptions();
        options.Mode = HtmlRenderMode.Paged;
        options.HonorCssPageRules = true;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        Assert.Equal(2, rendered.Pages.Count);
        HtmlRenderText[] following = rendered.Pages[1].Visuals.OfType<HtmlRenderText>()
            .Where(text => text.SemanticRole != "page-margin" && text.Y >= 200D).ToArray();
        Assert.Equal("Next inline continuation", string.Concat(following.Select(text => text.Text)));
        Assert.All(following, text => Assert.InRange(text.Y, 200D, 249D));
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(), text => text.SemanticRole == "page-margin" && text.Text == "Next inline continuation");
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.SemanticRole == "page-margin" && text.Text == "Next inline continuation");
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlTablePercentageHeights_PercentageRowspanFallbackIdentifiesTheEffectivePropertyAndRejectsStrictLoss() {
        string html = TableGeometrySource("<table style='height:200px;width:240px;margin:0;border-spacing:0'>"
            + "<tr><td id='span' rowspan='2' style='height:50%;padding:0;background:lime'>Span</td><td style='padding:0'>First</td></tr>"
            + "<tr><td style='padding:0'>Second</td></tr><tr><td colspan='2' style='padding:0'>Third</td></tr></table>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());
        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.TableValueUnsupported);

        Assert.Equal("td#span", diagnostic.Source);
        Assert.Contains("height=50%", diagnostic.Detail);
        Assert.Equal(OfficeConversionLossKind.Approximation, diagnostic.LossKind);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
        Assert.Equal(new[] { "Span", "First", "Second", "Third" }, rendered.Pages[0].Visuals.OfType<HtmlRenderText>().Select(item => item.Text));
    }

    [Fact]
    public void HtmlTablePercentageHeights_SharesExceedingTheGridKeepContentAndIdentifyTheApproximation() {
        string html = TableGeometrySource("<table style='height:200px;width:240px;margin:0;border-spacing:0'>"
            + "<tr><td id='summary' style='height:75%;padding:0;background:lime'>Summary</td></tr>"
            + "<tr><td id='detail' style='height:75%;padding:0;background:lime'>Detail</td></tr></table>"
            + "<div id='next' style='height:20px;background:red'>Next</div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());
        HtmlDiagnostic[] diagnostics = rendered.Diagnostics.Where(item => item.Code == HtmlRenderDiagnosticCodes.TableValueUnsupported).ToArray();

        Assert.Equal(new[] { "td#summary", "td#detail" }, diagnostics.Select(item => item.Source));
        Assert.All(diagnostics, item => Assert.Equal(OfficeConversionLossKind.Approximation, item.LossKind));
        Assert.All(diagnostics, item => Assert.Contains("height=75%", item.Detail));
        Assert.Equal(200D, TableGeometryShape(rendered, "div#next").Y, 3);
        Assert.Equal(new[] { "Summary", "Detail", "Next" }, rendered.Pages[0].Visuals.OfType<HtmlRenderText>().Select(item => item.Text));
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }
}

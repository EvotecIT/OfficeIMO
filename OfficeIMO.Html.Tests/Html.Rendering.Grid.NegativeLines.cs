using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(240D)]
    [InlineData(680D)]
    public void HtmlGrid_NegativeEndLineKeepsQueueSubgridRowsSeparate(double width) {
        string html = "<style>.queue{display:grid;grid-template-columns:22px minmax(0,1fr) 50px 60px;column-gap:8px}"
            + ".row{grid-column:1 / -1;display:grid;grid-template-columns:subgrid;padding:4px 0;background:#eee}"
            + ".title{grid-column:2}.metric{grid-column:3}.status{grid-column:4}</style>"
            + "<div class='queue'>"
            + string.Concat(Enumerable.Range(1, 3).Select(index => $"<div class='row' id='queue-{index}'><span>{index}</span>"
                + $"<span class='title' id='title-{index}' style='background:red'>Incident {index}</span>"
                + "<span class='metric'>12 min</span><span class='status'>Open</span></div>"))
            + "</div>";

        HtmlRenderDocument rendered = RenderGrid(html, width);
        HtmlRenderShape[] rows = Enumerable.Range(1, 3).Select(index => FindGridShape(rendered, "div#queue-" + index)).ToArray();
        Assert.All(rows, row => Assert.Equal(width, row.Width, 3));
        for (int index = 1; index < rows.Length; index++) {
            Assert.True(rows[index].Y >= rows[index - 1].Y + rows[index - 1].Height - 0.01D,
                $"Row {index}: previous Y={rows[index - 1].Y}, height={rows[index - 1].Height}; next Y={rows[index].Y}, height={rows[index].Height}");
        }
        Assert.All(Enumerable.Range(1, 3), index => {
            HtmlRenderShape title = FindGridShape(rendered, "span#title-" + index);
            Assert.Equal(30D, title.X, 3);
            Assert.True(title.Width > 70D);
        });
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GridValueUnsupported);
    }

    [Theory]
    [InlineData("grid")]
    [InlineData("inline-grid")]
    public void HtmlGrid_NegativeLinesStayAnchoredToExplicitTracksAfterImplicitExpansion(string display) {
        string html = $"<div style='display:{display};grid-template-columns:40px 60px;grid-template-rows:20px 30px;grid-auto-columns:50px;grid-auto-rows:25px'>"
            + "<span id='implicit' style='grid-column:3;grid-row:3;background:blue'>A</span>"
            + "<span id='last-explicit' style='grid-column:-2 / -1;grid-row:-2 / -1;background:red'>B</span></div>";

        HtmlRenderDocument rendered = RenderGrid(html, 200D);
        HtmlRenderShape last = FindGridShape(rendered, "span#last-explicit");
        Assert.Equal(40D, last.X, 3);
        Assert.Equal(20D, last.Y, 3);
        Assert.Equal(60D, last.Width, 3);
        Assert.Equal(30D, last.Height, 3);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GridValueUnsupported);
    }

    [Fact]
    public void HtmlGrid_NegativeLinesCreateLeadingTracksWithoutMovingNamedLinesOrAreas() {
        const string html = "<div style='position:relative;display:grid;width:130px;grid-template-columns:[left] 40px [middle] 60px [right];"
            + "grid-template-rows:[top] 20px [bottom];grid-template-areas:\"first second\";grid-auto-columns:20px 30px;grid-auto-rows:10px 15px'>"
            + "<span id='leading' style='grid-column:-4 / -3;grid-row:-3 / -2;background:blue'>A</span>"
            + "<span id='named-explicit' style='grid-column:left / middle;grid-row:top / bottom;background:green'>B</span>"
            + "<span id='positioned-area' style='position:absolute;grid-area:second;left:0;right:0;top:0;bottom:0;background:red'>C</span></div>";

        HtmlRenderDocument rendered = RenderGrid(html, 160D);
        HtmlRenderShape leading = FindGridShape(rendered, "span#leading");
        HtmlRenderShape explicitItem = FindGridShape(rendered, "span#named-explicit");
        HtmlRenderShape area = FindGridShape(rendered, "span#positioned-area");
        Assert.Equal(30D, leading.Width, 3);
        Assert.Equal(15D, leading.Height, 3);
        Assert.Equal(30D, explicitItem.X, 3);
        Assert.Equal(15D, explicitItem.Y, 3);
        Assert.Equal(40D, explicitItem.Width, 3);
        Assert.Equal(70D, area.X, 3);
        Assert.Equal(15D, area.Y, 3);
        Assert.Equal(60D, area.Width, 3);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GridValueUnsupported);
    }

    [Fact]
    public void HtmlGrid_NegativeLastLineUsesTheExplicitOriginWhenNoTracksAreDeclared() {
        const string html = "<div style='display:grid;width:50px;grid-auto-columns:20px 30px;grid-auto-rows:10px 15px'>"
            + "<span id='before-origin' style='grid-column:-2 / -1;grid-row:-2 / -1;background:blue'>A</span>"
            + "<span id='after-origin' style='grid-column:1;grid-row:1;background:red'>B</span></div>";
        HtmlRenderDocument rendered = RenderGrid(html, 100D);
        HtmlRenderShape before = FindGridShape(rendered, "span#before-origin");
        HtmlRenderShape after = FindGridShape(rendered, "span#after-origin");
        Assert.Equal(30D, before.Width, 3);
        Assert.Equal(15D, before.Height, 3);
        Assert.Equal(30D, after.X, 3);
        Assert.Equal(15D, after.Y, 3);
        Assert.Equal(20D, after.Width, 3);
        Assert.Equal(10D, after.Height, 3);
    }

    [Theory]
    [InlineData("4 / 2")]
    [InlineData("-1 / -3")]
    public void HtmlGrid_ReversedNumericEndpointsResolveTheSameArea(string placement) {
        HtmlRenderDocument rendered = RenderGrid("<div style='display:grid;grid-template-columns:20px 30px 40px;grid-template-rows:20px'>"
            + $"<span id='reversed' style='grid-column:{placement};background:red'>A</span></div>", 100D);
        HtmlRenderShape item = FindGridShape(rendered, "span#reversed");
        Assert.Equal(20D, item.X, 3);
        Assert.Equal(70D, item.Width, 3);
    }

    [Fact]
    public void HtmlGrid_PositionedNegativeLinesIgnoreTrailingImplicitTracks() {
        const string html = "<div style='position:relative;display:grid;width:150px;grid-template-columns:40px 60px;grid-template-rows:20px 30px;grid-auto-columns:50px;grid-auto-rows:25px'>"
            + "<span style='grid-column:3;grid-row:3'>A</span>"
            + "<span id='last-positioned' style='position:absolute;grid-column:-2 / -1;grid-row:-2 / -1;left:0;right:0;top:0;bottom:0;background:red'>B</span></div>";
        HtmlRenderDocument rendered = RenderGrid(html, 200D);
        HtmlRenderShape item = FindGridShape(rendered, "span#last-positioned");
        Assert.Equal(40D, item.X, 3);
        Assert.Equal(20D, item.Y, 3);
        Assert.Equal(60D, item.Width, 3);
        Assert.Equal(30D, item.Height, 3);
    }

    [Fact]
    public void HtmlGrid_RowSubgridResolvesNegativeLinesAgainstItsInheritedSpan() {
        const string html = "<div style='display:grid;width:100px;grid-template-rows:20px 30px'>"
            + "<div style='display:grid;grid-row:1 / -1;grid-template-rows:subgrid'>"
            + "<span id='subgrid-last-row' style='grid-row:-2 / -1;background:red'>A</span></div></div>";
        HtmlRenderDocument rendered = RenderGrid(html, 120D);
        HtmlRenderShape item = FindGridShape(rendered, "span#subgrid-last-row");
        Assert.Equal(20D, item.Y, 3);
        Assert.Equal(30D, item.Height, 3);
    }

    [Theory]
    [InlineData(false, "grid-area:target;grid-column:1 / -1", 80D)]
    [InlineData(true, "grid-area:target;grid-column:1 / -1", 80D)]
    [InlineData(false, "grid-area:target;grid-column:initial", 30D)]
    [InlineData(true, "grid-area:target;grid-column:initial", 30D)]
    [InlineData(false, "grid-column:1 / -1!important;grid-area:target", 80D)]
    [InlineData(true, "grid-column:1 / -1!important;grid-area:target", 80D)]
    [InlineData(false, "--placement:1 / -1;grid-area:target;grid-column:var(--placement)", 80D)]
    [InlineData(true, "--placement:1 / -1;grid-area:target;grid-column:var(--placement)", 80D)]
    [InlineData(false, "--placement:initial;grid-area:target;grid-column:var(--placement)", 30D)]
    [InlineData(true, "--placement:initial;grid-area:target;grid-column:var(--placement)", 30D)]
    [InlineData(false, "grid-area:target;grid-column:var(--missing)", 30D)]
    [InlineData(true, "grid-area:target;grid-column:var(--missing)", 30D)]
    [InlineData(false, "--placement:1 / 2 / 3;grid-area:target;grid-column:var(--placement)", 30D)]
    [InlineData(true, "--placement:1 / 2 / 3;grid-area:target;grid-column:var(--placement)", 30D)]
    public void HtmlGrid_NamedAreaHonorsIndependentColumnCascade(bool stylesheet, string declarations, double expectedWidth) {
        string html = (stylesheet ? "<style>#item{" + declarations + "}</style>" : "")
            + "<div style='display:grid;width:80px;grid-template-columns:30px 50px;grid-template-rows:20px 30px;grid-template-areas:\"first first\" \"left target\"'>"
            + "<span id='item' style='background:red;" + (stylesheet ? "" : declarations) + "'>A</span></div>";
        HtmlRenderDocument rendered = RenderGrid(html, 100D);
        HtmlRenderShape item = FindGridShape(rendered, "span#item");
        Assert.Equal(0D, item.X, 3);
        Assert.Equal(20D, item.Y, 3);
        Assert.Equal(expectedWidth, item.Width, 3);
        Assert.Equal(30D, item.Height, 3);
    }

    [Theory]
    [InlineData("-4 / -2", 0D, 40D)]
    [InlineData("2 / 5", 40D, 60D)]
    public void HtmlGrid_SubgridClampsBothEndpointsToInheritedColumns(string placement, double expectedX, double expectedWidth) {
        string html = "<div style='display:grid;width:100px;grid-template-columns:40px 60px'>"
            + "<div style='display:grid;grid-column:1 / -1;grid-template-columns:subgrid'>"
            + $"<span id='clamped' style='grid-column:{placement};background:red'>A</span></div></div>";
        HtmlRenderDocument rendered = RenderGrid(html, 120D);
        HtmlRenderShape item = FindGridShape(rendered, "span#clamped");
        Assert.Equal(expectedX, item.X, 3);
        Assert.Equal(expectedWidth, item.Width, 3);
    }

    [Theory]
    [InlineData(600D)]
    [InlineData(1000D)]
    public void HtmlGrid_QueueSubgridContributesEachChildToItsOwnAutomaticColumn(double width) {
        string html = "<style>.queue{display:grid;grid-template-columns:22px minmax(0,1fr) auto auto;column-gap:12px}"
            + ".row{display:grid;grid-column:1 / -1;grid-template-columns:subgrid;padding:4px 0;background:#eee}"
            + ".metric,.status{white-space:nowrap}</style><div class='queue'>"
            + string.Concat(Enumerable.Range(1, 3).Select(index => $"<div class='row' id='queue-{index}'><span>{index}</span>"
                + $"<div id='title-{index}' style='background:red'>Directory health on corp-dc004.corp.example"
                + "<p>Not up for 2 h of the last 1 d. Timed out after 5000 ms waiting for corp-dc004.corp.example.</p></div>"
                + $"<span class='metric' id='metric-{index}' style='background:blue'>12 checks</span><span class='status'>Down</span></div>"))
            + "</div>";
        HtmlRenderDocument rendered = RenderGrid(html, width);
        HtmlRenderShape[] rows = Enumerable.Range(1, 3).Select(index => FindGridShape(rendered, "div#queue-" + index)).ToArray();
        Assert.All(Enumerable.Range(1, 3), index => {
            HtmlRenderShape title = FindGridShape(rendered, "div#title-" + index);
            HtmlRenderShape metric = FindGridShape(rendered, "span#metric-" + index);
            Assert.True(title.Width > width - 250D, $"Title column width was {title.Width} for {width} available pixels.");
            Assert.InRange(metric.Width, 40D, 150D);
            Assert.True(metric.X >= title.X + title.Width);
        });
        for (int index = 1; index < rows.Length; index++) Assert.True(rows[index].Y >= rows[index - 1].Y + rows[index - 1].Height);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GridValueUnsupported);
    }

    [Theory]
    [InlineData(false, "grid-column:0 / -1")]
    [InlineData(true, "grid-column:0 / -1")]
    [InlineData(false, "grid-column:span -2")]
    [InlineData(true, "grid-column:span -2")]
    [InlineData(false, "grid-column:initial / -1")]
    [InlineData(true, "grid-column:initial / -1")]
    [InlineData(false, "grid-column-start:0")]
    [InlineData(true, "grid-column-start:0")]
    public void HtmlGrid_InvalidPlacementDeclarationPreservesEarlierValidColumn(bool stylesheet, string invalid) {
        string declarations = "grid-column:1 / -1;" + invalid;
        string html = (stylesheet ? "<style>#valid{" + declarations + "}</style>" : "")
            + "<div style='display:grid;width:100px;grid-template-columns:40px 60px'>"
            + "<span id='valid' style='background:red;" + (stylesheet ? "" : declarations) + "'>A</span></div>";
        HtmlRenderShape item = FindGridShape(RenderGrid(html, 120D), "span#valid");
        Assert.Equal(0D, item.X, 3);
        Assert.Equal(100D, item.Width, 3);
    }

    [Theory]
    [InlineData("<span style='grid-column:-2147483648'>A</span>")]
    [InlineData("<span style='grid-column:-4 / -3'>A</span><span style='grid-column:4'>B</span>")]
    public void HtmlGrid_BoundsNegativeLinesAndCombinedImplicitExpansion(string items) {
        HtmlDomLimitException error = Assert.Throws<HtmlDomLimitException>(() => HtmlRenderTestDriver.Render(
            "<div style='display:grid;grid-template-columns:20px 30px'>" + items + "</div>", new HtmlRenderOptions {MaxGridTracks=4}));
        Assert.Equal(HtmlRenderDiagnosticCodes.GridTrackLimitExceeded, error.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxGridTracks), error.LimitSource);
    }
}

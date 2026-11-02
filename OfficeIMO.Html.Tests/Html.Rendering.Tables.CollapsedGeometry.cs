using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("collapse", 31D, 63D, 0.5D)]
    [InlineData("separate", 32D, 64D, 0D)]
    public void HtmlCollapsedGeometry_TwoLineRowsUseTheWinningHalfBorders(
        string model, double rowHeight, double tableHeight, double perimeterOffset) {
        string html = "<table id='grid' style='width:280px;border-collapse:" + model + "'><tbody>"
            + "<tr><td id='a' style='border:1px solid black;padding:5px'>A0<br>B0</td></tr>"
            + "<tr><td id='b' style='border:1px solid black;padding:5px'>A1<br>B1</td></tr></tbody></table>"
            + CollapsedGeometryFollowing;
        HtmlRenderDocument rendered = RenderCollapsedGeometry(html);
        HtmlRenderShape table = TableGeometryShape(rendered, "table#grid");
        HtmlRenderShape first = TableGeometryShape(rendered, "td#a");
        HtmlRenderShape second = TableGeometryShape(rendered, "td#b");
        HtmlRenderText firstText = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "A0");
        HtmlRenderText secondText = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "A1");

        Assert.Equal(280D, table.Width, 3);
        Assert.Equal(tableHeight, table.Height, 3);
        Assert.Equal(perimeterOffset, first.X - table.X, 3);
        Assert.Equal(perimeterOffset, first.Y - table.Y, 3);
        Assert.Equal(280D - perimeterOffset * 2D, first.Width, 3);
        Assert.Equal(rowHeight, first.Height, 3);
        Assert.Equal(rowHeight, second.Height, 3);
        Assert.Equal(rowHeight, second.Y - first.Y, 3);
        Assert.Equal(rowHeight, secondText.Y - firstText.Y, 3);
        Assert.Equal(tableHeight, TableGeometryShape(rendered, "div#after").Y, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("10px solid red", "2px solid blue", "10px solid red")]
    [InlineData("6px dashed red", "6px solid blue", "6px solid blue")]
    [InlineData("8px none red", "6px solid blue", "6px solid blue")]
    public void HtmlCollapsedGeometry_NeighborConflictsMatchTheirResolvedContentBoxes(
        string left, string right, string winner) {
        HtmlRenderDocument conflict = RenderCollapsedGeometry(CollapsedGeometryNeighbors(left, right));
        HtmlRenderDocument resolved = RenderCollapsedGeometry(CollapsedGeometryNeighbors(winner, winner));

        AssertCollapsedGeometryEquivalent(resolved, conflict, "table#grid", "td#a", "td#b",
            "div#a-probe", "div#b-probe", "div#after");
    }

    [Fact]
    public void HtmlCollapsedGeometry_HiddenReclaimsSpaceButTransparentWinningBordersRetainIt() {
        HtmlRenderDocument hidden = RenderCollapsedGeometry(CollapsedGeometryNeighbors("10px solid red", "4px hidden blue"));
        HtmlRenderDocument none = RenderCollapsedGeometry(CollapsedGeometryNeighbors("none", "none"));
        HtmlRenderDocument transparent = RenderCollapsedGeometry(CollapsedGeometryNeighbors("10px solid transparent", "2px solid blue"));
        HtmlRenderDocument visible = RenderCollapsedGeometry(CollapsedGeometryNeighbors("10px solid red", "10px solid red"));

        AssertCollapsedGeometryEquivalent(none, hidden, "table#grid", "td#a", "td#b",
            "div#a-probe", "div#b-probe", "div#after");
        AssertCollapsedGeometryEquivalent(visible, transparent, "table#grid", "td#a", "td#b",
            "div#a-probe", "div#b-probe", "div#after");
        Assert.True(TableGeometryShape(hidden, "div#a-probe").Width > TableGeometryShape(transparent, "div#a-probe").Width);
        Assert.True(TableGeometryShape(hidden, "div#b-probe").Width > TableGeometryShape(transparent, "div#b-probe").Width);
    }

    [Fact]
    public void HtmlCollapsedGeometry_TablePaddingDoesNotDisplaceTheGridOrFollowingContent() {
        string Table(int padding) => "<table id='grid' style='width:200px;table-layout:fixed;border-collapse:collapse;"
            + "border:4px solid purple;padding:" + padding + "px'><tr>"
            + "<td id='a'><div id='a-probe' class='probe'>Left</div></td>"
            + "<td id='b'><div id='b-probe' class='probe'>Right</div></td></tr></table>" + CollapsedGeometryFollowing;
        HtmlRenderDocument ordinary = RenderCollapsedGeometry(Table(0));
        HtmlRenderDocument padded = RenderCollapsedGeometry(Table(17));

        AssertCollapsedGeometryEquivalent(ordinary, padded, "table#grid", "td#a", "td#b",
            "div#a-probe", "div#b-probe", "div#after");
        HtmlRenderShape cell = TableGeometryShape(padded, "td#a");
        Assert.True(TableGeometryShape(padded, "div#a-probe").X > cell.X);
    }

    [Fact]
    public void HtmlCollapsedGeometry_AutoIntrinsicWidthCountsTheWinningInsetsOnce() {
        string Table(string neighbor) => "<table id='grid' style='width:auto;table-layout:auto;border-collapse:collapse'><tr>"
            + "<td id='a' style='border-right:10px solid red'><div id='a-probe' class='probe' style='width:40px'>Left</div></td>"
            + "<td id='b' style='border-left:" + neighbor + "'><div id='b-probe' class='probe' style='width:60px'>Right</div></td>"
            + "</tr></table>" + CollapsedGeometryFollowing;
        HtmlRenderDocument conflict = RenderCollapsedGeometry(Table("2px solid blue"));
        HtmlRenderDocument resolved = RenderCollapsedGeometry(Table("10px solid red"));

        AssertCollapsedGeometryEquivalent(resolved, conflict, "table#grid", "td#a", "td#b",
            "div#a-probe", "div#b-probe", "div#after");
        Assert.True(TableGeometryShape(conflict, "table#grid").Width < TableGeometryOptions().ViewportWidth);
        Assert.Equal(40D, TableGeometryShape(conflict, "div#a-probe").Width, 3);
        Assert.Equal(60D, TableGeometryShape(conflict, "div#b-probe").Width, 3);
    }

    [Fact]
    public void HtmlCollapsedGeometry_FixedColumnWrappingUsesTheResolvedContentWidth() {
        const string words = "alpha beta gamma delta";
        string Table(string neighbor) => "<table id='grid' style='width:100px;table-layout:fixed;border-collapse:collapse'><tr>"
            + "<td id='a' style='border-right:20px solid red'><div id='a-probe' style='background:lime'>" + words + "</div></td>"
            + "<td id='b' style='border-left:" + neighbor + "'><div id='b-probe' style='background:lime'>Peer</div></td>"
            + "</tr></table>" + CollapsedGeometryFollowing;
        HtmlRenderDocument conflict = RenderCollapsedGeometry(Table("2px solid blue"));
        HtmlRenderDocument resolved = RenderCollapsedGeometry(Table("20px solid red"));

        AssertCollapsedGeometryEquivalent(resolved, conflict, "table#grid", "td#a", "td#b",
            "div#a-probe", "div#b-probe", "div#after");
        Assert.True(TableGeometryShape(conflict, "div#a-probe").Height > 10D);
        string actual = string.Concat(conflict.Pages[0].Visuals.OfType<HtmlRenderText>()
            .Where(text => text.Text != "Peer" && text.Text != "After").Select(text => text.Text));
        Assert.Equal(words.Replace(" ", string.Empty), string.Concat(actual.Where(character => !char.IsWhiteSpace(character))));
    }

    [Fact]
    public void HtmlCollapsedGeometry_RowTrackGroupAndTableWinnersSupplyTheCellInsets() {
        string Cell(string id, string style) => "<td id='" + id + "' style='" + style
            + "'><div id='" + id + "-probe' class='probe'>" + id.ToUpperInvariant() + "</div></td>";
        string conflict = "<table id='grid' style='width:200px;table-layout:fixed;border-collapse:collapse;border:6px solid purple'>"
            + "<colgroup style='border:6px solid orange'><col style='border-right:6px solid blue'></colgroup><colgroup><col></colgroup>"
            + "<tbody style='border-top:6px solid green;border-bottom:6px solid green'>"
            + "<tr style='border-top:6px solid red;border-bottom:6px solid red'>"
            + Cell("a", "border:none;border-top:6px solid black;border-right:6px solid black") + Cell("b", "border:none")
            + "</tr><tr>" + Cell("c", "border:none") + Cell("d", "border:none") + "</tr></tbody></table>" + CollapsedGeometryFollowing;
        string resolved = "<table id='grid' style='width:200px;table-layout:fixed;border-collapse:collapse'><tr>"
            + Cell("a", "border-top:6px solid black;border-right:6px solid black;border-bottom:6px solid red;border-left:6px solid orange")
            + Cell("b", "border-top:6px solid red;border-right:6px solid purple;border-bottom:6px solid red;border-left:6px solid black")
            + "</tr><tr>"
            + Cell("c", "border-top:6px solid red;border-right:6px solid blue;border-bottom:6px solid green;border-left:6px solid orange")
            + Cell("d", "border-top:6px solid red;border-right:6px solid purple;border-bottom:6px solid green;border-left:6px solid blue")
            + "</tr></table>" + CollapsedGeometryFollowing;

        AssertCollapsedGeometryEquivalent(RenderCollapsedGeometry(resolved), RenderCollapsedGeometry(conflict),
            "table#grid", "td#a", "td#b", "td#c", "td#d", "div#a-probe", "div#b-probe", "div#c-probe", "div#d-probe", "div#after");
    }

    [Theory]
    [InlineData("auto", "ltr")]
    [InlineData("fixed", "ltr")]
    [InlineData("fixed", "rtl")]
    public void HtmlCollapsedGeometry_SpansKeepSegmentedWinnersAndLogicalColumnPlacement(string layout, string direction) {
        string leading = direction == "rtl" ? "left" : "right";
        string trailing = direction == "rtl" ? "right" : "left";
        string Cell(string id, string span, string style, string text) => "<td id='" + id + "' " + span
            + " style='" + style + "'><div id='" + id + "-probe' class='probe'>" + text + "</div></td>";
        string Table(bool resolved) => "<table id='grid' dir='" + direction + "' style='border-collapse:collapse;table-layout:" + layout
            + ";width:" + (layout == "auto" ? "auto" : "280px") + "'><tr>"
            + Cell("a", "rowspan='2'", "border-" + leading + ":" + (resolved ? "none" : "2px solid blue"), "Span")
            + Cell("b", "colspan='2'", "border-" + trailing + ":10px solid red;border-bottom:8px solid purple", "Wide span content")
            + "</tr><tr>"
            + Cell("c", string.Empty, "border-" + trailing + ":6px solid green;border-top:" + (resolved ? "8px solid purple" : "2px solid black"), "Lower")
            + Cell("d", string.Empty, "border-top:" + (resolved ? "8px solid purple" : "4px solid gray"), "Cell")
            + "</tr></table>" + CollapsedGeometryFollowing;
        HtmlRenderDocument conflict = RenderCollapsedGeometry(Table(false));
        HtmlRenderDocument resolved = RenderCollapsedGeometry(Table(true));

        AssertCollapsedGeometryEquivalent(resolved, conflict, "table#grid", "td#a", "td#b", "td#c", "td#d",
            "div#a-probe", "div#b-probe", "div#c-probe", "div#d-probe", "div#after");
        HtmlRenderShape first = TableGeometryShape(conflict, "td#a");
        HtmlRenderShape upper = TableGeometryShape(conflict, "td#b");
        HtmlRenderShape lower = TableGeometryShape(conflict, "td#c");
        Assert.True(direction == "rtl" ? first.X > lower.X : first.X < lower.X);
        Assert.True(first.Height > upper.Height);
        Assert.True(upper.Width > lower.Width);
        Assert.True(lower.Y > upper.Y);
    }

    [Fact]
    public void HtmlCollapsedGeometry_PercentageDescendantsDeductTheFinalWinningInsetsOnce() {
        string Table(string neighbor) => "<table id='grid' style='width:200px;height:120px;table-layout:fixed;border-collapse:collapse'><tr>"
            + "<td id='a' style='height:50%;padding:4px;border:8px solid black;border-bottom:10px solid red'>"
            + "<div id='a-probe' style='height:100%;background:lime'>Fill</div></td></tr><tr>"
            + "<td id='b' style='height:50%;padding:4px;border:8px solid black;border-top:" + neighbor + "'>"
            + "<div id='b-probe' style='height:100%;background:lime'>Due</div></td></tr></table>" + CollapsedGeometryFollowing;
        HtmlRenderDocument conflict = RenderCollapsedGeometry(Table("2px solid blue"));
        HtmlRenderDocument resolved = RenderCollapsedGeometry(Table("10px solid red"));

        AssertCollapsedGeometryEquivalent(resolved, conflict, "table#grid", "td#a", "td#b",
            "div#a-probe", "div#b-probe", "div#after");
        HtmlRenderShape first = TableGeometryShape(conflict, "td#a");
        HtmlRenderShape second = TableGeometryShape(conflict, "td#b");
        HtmlRenderShape fill = TableGeometryShape(conflict, "div#a-probe");
        HtmlRenderShape detail = TableGeometryShape(conflict, "div#b-probe");
        Assert.Equal(first.Y + 8D, fill.Y, 3);
        Assert.Equal(second.Y + 9D, detail.Y, 3);
        Assert.Equal(first.Height - 17D, fill.Height, 3);
        Assert.Equal(second.Height - 17D, detail.Height, 3);
        HtmlRenderShape table = TableGeometryShape(conflict, "table#grid");
        Assert.Equal(table.Y + table.Height, TableGeometryShape(conflict, "div#after").Y, 3);
    }

    [Fact]
    public void HtmlCollapsedGeometry_ContinuationKeepsUnequalGroupBordersAndEveryBodyRowOnce() {
        string rows = string.Concat(Enumerable.Range(1, 8).Select(index => "<tr style='break-inside:avoid'><td id='row" + index
            + "' style='height:24px;padding:2px;border-top:" + (index % 2 == 0 ? "6px solid blue" : "2px solid blue")
            + "'><div class='probe' style='height:12px'>Body" + index + "</div></td></tr>"));
        string html = "<table id='grid' style='width:180px;table-layout:fixed;border-collapse:collapse'>"
            + "<thead><tr><th id='head' style='padding:2px;border-bottom:10px solid red'>"
            + "<div id='head-probe' class='probe' style='height:12px'>Header</div></th></tr></thead><tbody>" + rows
            + "</tbody><tfoot><tr><td id='foot' style='padding:2px;border-top:8px solid green'>"
            + "<div id='foot-probe' class='probe' style='height:12px'>Footer</div></td></tr></tfoot></table>" + CollapsedGeometryFollowing;
        HtmlRenderOptions options = CollapsedGeometryOptions();
        options.Mode = HtmlRenderMode.Paged;
        options.PageSize = new OfficePageSize(global::OfficeIMO.OfficeLength.Inches(220D / 96D), global::OfficeIMO.OfficeLength.Inches(120D / 96D));
        options.HonorCssPageRules = false;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(CollapsedGeometrySource(html), options);
        HtmlRenderPage[] bodyPages = rendered.Pages.Where(page => CollapsedGeometryPageVisuals(page).OfType<HtmlRenderText>()
            .Any(text => text.Text.StartsWith("Body", StringComparison.Ordinal))).ToArray();
        HtmlRenderText[] text = rendered.Pages.SelectMany(CollapsedGeometryPageVisuals).OfType<HtmlRenderText>().ToArray();

        Assert.True(bodyPages.Length > 1);
        foreach (int index in Enumerable.Range(1, 8)) Assert.Single(text, item => item.Text == "Body" + index);
        Assert.Single(text, item => item.Text == "After");
        double? headerHeight = null;
        double? footerHeight = null;
        foreach (HtmlRenderPage page in bodyPages) {
            HtmlRenderVisual[] visuals = CollapsedGeometryPageVisuals(page).ToArray();
            Assert.Single(visuals.OfType<HtmlRenderText>(), item => item.Text == "Header");
            Assert.Single(visuals.OfType<HtmlRenderText>(), item => item.Text == "Footer");
            HtmlRenderShape head = Assert.Single(visuals.OfType<HtmlRenderShape>(), item => item.Source == "th#head" && item.Shape.FillColor.HasValue);
            HtmlRenderShape foot = Assert.Single(visuals.OfType<HtmlRenderShape>(), item => item.Source == "td#foot" && item.Shape.FillColor.HasValue);
            headerHeight ??= head.Height;
            footerHeight ??= foot.Height;
            Assert.Equal(headerHeight.Value, head.Height, 3);
            Assert.Equal(footerHeight.Value, foot.Height, 3);
            HtmlRenderShape[] cells = visuals.OfType<HtmlRenderShape>().Where(item => item.Source != null
                && item.Source.StartsWith("td#row", StringComparison.Ordinal) && item.Shape.FillColor.HasValue).ToArray();
            Assert.NotEmpty(cells);
            Assert.All(cells, cell => {
                Assert.True(cell.Y >= head.Y + head.Height - 0.001D);
                Assert.True(cell.Y + cell.Height <= foot.Y + 0.001D);
            });
        }

        HtmlRenderPage followingPage = Assert.Single(rendered.Pages, page => CollapsedGeometryPageVisuals(page).OfType<HtmlRenderText>()
            .Any(item => item.Text == "After"));
        Assert.True(followingPage.PageNumber >= bodyPages[bodyPages.Length - 1].PageNumber);
        if (followingPage.PageNumber == bodyPages[bodyPages.Length - 1].PageNumber) {
            HtmlRenderShape foot = Assert.Single(CollapsedGeometryPageVisuals(followingPage).OfType<HtmlRenderShape>(),
                item => item.Source == "td#foot" && item.Shape.FillColor.HasValue);
            HtmlRenderShape following = Assert.Single(CollapsedGeometryPageVisuals(followingPage).OfType<HtmlRenderShape>(),
                item => item.Source == "div#after" && item.Shape.FillColor.HasValue);
            Assert.True(following.Y >= foot.Y + foot.Height - 0.001D);
        }
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlCollapsedGeometry_RepeatedOutsideHalfBordersReserveSpaceBeforeFollowingContent() {
        string rows = string.Concat(Enumerable.Range(1, 12).Select(index =>
            "<tr style='break-inside:avoid'><td><div class='probe'>Body" + index.ToString("D2") + "<br>Detail</div></td></tr>"));
        string html = "<style>@page{size:320px 220px;margin:10px}td,th{padding:4px}</style>"
            + "<table id='grid' style='width:280px;table-layout:fixed;border-collapse:collapse'><thead><tr>"
            + "<th style='border-bottom:10px solid red'><div class='probe'>Header<br>Head detail</div></th></tr></thead>"
            + "<tbody>" + rows + "</tbody><tfoot><tr><td style='border-top:8px solid green'>"
            + "<div class='probe'>Footer<br>Foot detail</div></td></tr></tfoot></table>" + CollapsedGeometryFollowing;
        HtmlRenderOptions options = CollapsedGeometryOptions();
        options.Mode = HtmlRenderMode.Paged;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(CollapsedGeometrySource(html), options);

        Assert.Equal(4, rendered.Pages.Count);
        for (int index = 0; index < 3; index++) {
            HtmlRenderText[] text = CollapsedGeometryPageVisuals(rendered.Pages[index]).OfType<HtmlRenderText>().ToArray();
            Assert.Equal(15D, Assert.Single(text, item => item.Text == "Header").Y, 3);
            Assert.Single(text, item => item.Text == "Footer");
            Assert.Equal(Enumerable.Range(index * 4 + 1, 4).Select(number => "Body" + number.ToString("D2")),
                text.Where(item => item.Text.StartsWith("Body", StringComparison.Ordinal)).Select(item => item.Text));
        }
        Assert.Equal(new[] { "After" }, CollapsedGeometryPageVisuals(rendered.Pages[3]).OfType<HtmlRenderText>().Select(item => item.Text));
        IReadOnlyList<OfficeImageExportResult> png = HtmlConversionDocument.Parse(CollapsedGeometrySource(html))
            .ExportImages(OfficeImageExportFormat.Png, options);
        Assert.True(OfficePngReader.TryDecode(png[2].Bytes, out OfficeRasterImage? raster));
        Assert.Equal(OfficeColor.Red, raster!.GetPixel(285, 48));
        Assert.Equal(OfficeColor.Black, raster.GetPixel(285, 44));
        Assert.Equal(OfficeColor.Black, raster.GetPixel(285, 45));
        Assert.Equal(OfficeColor.White, raster.GetPixel(7, 44));
        Assert.Equal(OfficeColor.White, raster.GetPixel(292, 44));
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlCollapsedGeometry_RowAndColumnRulesDoNotCrossSpanningCellInteriors() {
        const string body = "<table style='width:280px;table-layout:fixed;border-collapse:collapse'>"
            + "<col><col style='border-right:12px solid blue'><col><tbody>"
            + "<tr style='border-bottom:12px solid red'><td rowspan='2' style='background:white'>"
            + "<div class='probe'>Span</div></td><td colspan='2'><div class='probe'>Wide</div></td></tr>"
            + "<tr><td><div class='probe'>Lower</div></td><td><div class='probe'>Cell</div></td></tr></tbody></table>";
        Assert.True(OfficePngReader.TryDecode(HtmlConversionDocument.Parse(CollapsedGeometrySource(body))
            .ToPng(CollapsedGeometryOptions()), out OfficeRasterImage? raster));

        Assert.Equal(OfficeColor.White, raster!.GetPixel(40, 34));
        Assert.Equal(OfficeColor.Lime, raster.GetPixel(190, 15));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(250, 34));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(190, 55));
    }

    [Theory]
    [InlineData("fixed", "200px")]
    [InlineData("auto", "auto")]
    public void HtmlCollapsedGeometry_ColumnHintsExcludeColumnPaddingAndKeepBorderWinners(string layout, string width) {
        string Table(bool columns) => "<table id='grid' style='table-layout:" + layout + ";width:" + width
            + ";border-collapse:collapse'><col style='width:80px;" + (columns ? "padding:17px;border:10px solid red" : "")
            + "'><col style='width:120px;" + (columns ? "padding:17px;border:10px solid blue" : "") + "'><tbody><tr>"
            + "<td id='a' style='" + (columns ? "" : "border:10px solid red") + "'><div id='a-probe' class='probe'>Left</div></td>"
            + "<td id='b' style='" + (columns ? "" : "border:10px solid blue") + "'><div id='b-probe' class='probe'>Right</div></td>"
            + "</tr></tbody></table>" + CollapsedGeometryFollowing;
        HtmlRenderDocument actual = RenderCollapsedGeometry(Table(true));
        AssertCollapsedGeometryEquivalent(RenderCollapsedGeometry(Table(false)), actual,
            "table#grid", "td#a", "td#b", "div#a-probe", "div#b-probe", "div#after");
        Assert.Equal(210D, TableGeometryShape(actual, "table#grid").Width, 3);
        Assert.Equal(80D, TableGeometryShape(actual, "td#a").Width, 3);
        Assert.Equal(120D, TableGeometryShape(actual, "td#b").Width, 3);
    }

    [Fact]
    public void HtmlCollapsedGeometry_FirstRowContentWidthHintUsesWinningInsets() {
        string Table(string neighbor) => "<table id='grid' style='width:200px;table-layout:fixed;border-collapse:collapse'><tr>"
            + "<td id='a' style='width:80px;border-right:10px solid red'><div id='a-probe' class='probe'>Left</div></td>"
            + "<td id='b' style='border-left:" + neighbor + "'><div id='b-probe' class='probe'>Right</div></td></tr></table>"
            + CollapsedGeometryFollowing;
        AssertCollapsedGeometryEquivalent(RenderCollapsedGeometry(Table("10px solid red")),
            RenderCollapsedGeometry(Table("2px solid blue")), "table#grid", "td#a", "td#b", "div#a-probe", "div#b-probe", "div#after");
    }

    [Fact]
    public void HtmlCollapsedGeometry_DoublePaintStaysInsideTheWinningHalfWidthBand() {
        const string body = "<table id='grid' style='width:280px;table-layout:fixed;border-collapse:collapse'><tr>"
            + "<td style='border:9px double black;padding:4px;background:white'><div class='probe'>Cell</div></td></tr></table>"
            + "<div id='after' style='height:10px;background:orange'>After</div>";
        string source = CollapsedGeometrySource(body);
        HtmlRenderOptions options = CollapsedGeometryOptions();
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(source, options);
        Assert.Equal(46D, TableGeometryShape(rendered, "table#grid").Height, 3);
        Assert.True(OfficePngReader.TryDecode(HtmlConversionDocument.Parse(source).ToPng(options), out OfficeRasterImage? raster));
        foreach (int y in new[] { 1, 7, 38, 44 }) Assert.Equal(OfficeColor.Black, raster!.GetPixel(200, y));
        Assert.Equal(OfficeColor.White, raster!.GetPixel(200, 4));
        Assert.Equal(OfficeColor.Orange, raster.GetPixel(200, 46));
    }

    [Theory]
    [InlineData("double", 9, 48, 54)]
    [InlineData("groove", 8, 48, 52)]
    [InlineData("ridge", 8, 48, 52)]
    public void HtmlCollapsedGeometry_MultiStrokeRulesKeepBothComponentsAtSafePageCuts(string style, int width, int upper, int lower) {
        string border = "border:" + width + "px " + style + " black;padding:4px;background:white";
        string rows = string.Concat(Enumerable.Range(1, 12).Select(index => "<tr style='break-inside:avoid'><td style='"
            + border + "'><div class='probe'>Body" + index + "<br>Detail</div></td></tr>"));
        string body = "<style>@page{size:320px 220px;margin:10px}</style><div style='height:66px'></div>"
            + "<table style='width:280px;table-layout:fixed;border-collapse:collapse'><thead><tr><th style='" + border
            + ";border-color:red'><div class='probe'>Header<br>Detail</div></th></tr></thead><tbody>" + rows
            + "</tbody><tfoot><tr><td style='" + border + ";border-color:blue'><div class='probe'>Footer<br>Detail</div>"
            + "</td></tr></tfoot></table>";
        HtmlRenderOptions options = CollapsedGeometryOptions();
        options.Mode = HtmlRenderMode.Paged;
        IReadOnlyList<OfficeImageExportResult> pages = HtmlConversionDocument.Parse(CollapsedGeometrySource(body))
            .ExportImages(OfficeImageExportFormat.Png, options);
        Assert.True(pages.Count >= 3);
        Assert.True(OfficePngReader.TryDecode(pages[2].Bytes, out OfficeRasterImage? raster));
        Assert.NotEqual(OfficeColor.White, raster!.GetPixel(270, upper));
        Assert.NotEqual(OfficeColor.White, raster.GetPixel(270, lower));
        if (style == "double") {
            Assert.Equal(OfficeColor.Black, raster.GetPixel(270, upper));
            Assert.Equal(OfficeColor.Black, raster.GetPixel(270, lower));
            Assert.Equal(OfficeColor.White, raster.GetPixel(270, 51));
        } else {
            Assert.NotEqual(raster.GetPixel(270, upper), raster.GetPixel(270, lower));
        }
    }

    [Fact]
    public void HtmlCollapsedGeometry_PageHeightContinuationPreservesSourceCellInsets() {
        const string body = "<style>@page{size:320px 200px;margin:10px}@page:first{size:320px 190px}"
            + "td,th{padding:4px}tr{break-inside:avoid}</style>"
            + "<div id='spacer' style='height:90px'></div>"
            + "<table id='grid' style='width:280px;table-layout:fixed;border-collapse:collapse'><thead><tr>"
            + "<th style='border-bottom:10px solid red'><div class='probe'>Header<br>Detail</div></th>"
            + "</tr></thead><tbody><tr><td id='first'><div class='probe'>Body01<br>Detail</div></td></tr>"
            + "<tr><td id='resumed'><div id='resumed-probe' class='probe'>Body02<br>Detail</div></td></tr>"
            + "<tr><td id='last'><div class='probe'>Body03<br>Detail</div></td></tr></tbody></table>"
            + CollapsedGeometryFollowing;
        HtmlRenderOptions options = CollapsedGeometryOptions();
        options.Mode = HtmlRenderMode.Paged;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(CollapsedGeometrySource(body), options);

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Equal(190D, rendered.Pages[0].Height, 3);
        Assert.Equal(200D, rendered.Pages[1].Height, 3);
        Assert.Contains(CollapsedGeometryPageVisuals(rendered.Pages[1]).OfType<HtmlRenderText>(),
            text => text.Text == "Body02");
        Assert.All(rendered.Pages, page => Assert.Single(CollapsedGeometryPageVisuals(page).OfType<HtmlRenderText>(),
            text => text.Text == "Header"));
        HtmlRenderShape first = TableGeometryShape(rendered, "td#first");
        HtmlRenderShape resumed = TableGeometryShape(rendered, "td#resumed");
        HtmlRenderShape probe = TableGeometryShape(rendered, "div#resumed-probe");
        Assert.Equal(34D, first.Height, 3);
        Assert.Equal(30D, resumed.Height, 3);
        Assert.Equal(5D, probe.Y - resumed.Y, 3);
        Assert.Equal(30D, TableGeometryShape(rendered, "td#last").Y - resumed.Y, 3);
        HtmlRenderText[] text = rendered.Pages.SelectMany(CollapsedGeometryPageVisuals).OfType<HtmlRenderText>().ToArray();
        foreach (string label in new[] { "Body01", "Body02", "Body03", "After" }) {
            Assert.Single(text, item => item.Text == label);
        }
        IReadOnlyList<OfficeImageExportResult> pages = HtmlConversionDocument.Parse(CollapsedGeometrySource(body))
            .ExportImages(OfficeImageExportFormat.Png, options);
        Assert.True(OfficePngReader.TryDecode(pages[1].Bytes, out OfficeRasterImage? raster));
        Assert.Equal(OfficeColor.Red, raster!.GetPixel(200, 43));
        Assert.Equal(OfficeColor.Black, raster.GetPixel(200, 44));
        Assert.Equal(OfficeColor.Black, raster.GetPixel(200, 45));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(200, 46));
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlCollapsedGeometry_ResumedBodyBorderDoesNotBecomeAFutureHeaderBorder() {
        string rows = string.Concat(Enumerable.Range(1, 12).Select(index => "<tr><td"
            + (index == 2 ? " style='border-top:20px solid blue'" : string.Empty)
            + "><div class='probe'>Body" + index.ToString("D2") + "<br>Detail</div></td></tr>"));
        string body = "<style>@page{size:320px 200px;margin:10px}@page:first{size:320px 190px}"
            + "td,th{padding:4px}tr{break-inside:avoid}</style><div style='height:90px'></div>"
            + "<table style='width:280px;table-layout:fixed;border-collapse:collapse'><thead><tr>"
            + "<th style='border-bottom:10px solid red'><div class='probe'>Header<br>Detail</div></th>"
            + "</tr></thead><tbody>" + rows + "</tbody></table>" + CollapsedGeometryFollowing;
        string source = CollapsedGeometrySource(body);
        HtmlRenderOptions options = CollapsedGeometryOptions();
        options.Mode = HtmlRenderMode.Paged;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(source, options);
        Assert.Equal(4, rendered.Pages.Count);
        HtmlRenderText[] text = rendered.Pages.SelectMany(CollapsedGeometryPageVisuals).OfType<HtmlRenderText>().ToArray();
        for (int index = 1; index <= 12; index++) {
            Assert.Single(text, item => item.Text == "Body" + index.ToString("D2"));
        }
        IReadOnlyList<OfficeImageExportResult> pages = HtmlConversionDocument.Parse(source)
            .ExportImages(OfficeImageExportFormat.Png, options);
        Assert.True(OfficePngReader.TryDecode(pages[0].Bytes, out OfficeRasterImage? firstPage));
        Assert.Equal(OfficeColor.Blue, firstPage!.GetPixel(200, 172));
        Assert.Equal(OfficeColor.White, firstPage.GetPixel(200, 182));
        for (int index = 2; index < pages.Count; index++) {
            Assert.True(OfficePngReader.TryDecode(pages[index].Bytes, out OfficeRasterImage? raster));
            Assert.Equal(OfficeColor.Red, raster!.GetPixel(200, 43));
            Assert.Equal(OfficeColor.Black, raster.GetPixel(200, 44));
            Assert.Equal(OfficeColor.Red, raster.GetPixel(200, 46));
        }
        rendered.RequireNoLoss();
    }

    private const string CollapsedGeometryFollowing = "<div id='after' style='height:10px;background:red'>After</div>";

    private static string CollapsedGeometrySource(string body) => TableGeometrySource("<style>body{font:9px/10px Witness}"
        + "table{margin:0;border-spacing:0;background:white}td,th{border:2px solid black;padding:3px;vertical-align:top;"
        + "text-align:left;font-weight:400;background:#eee}div{margin:0}.probe{height:20px;background:lime}</style>" + body);

    private static HtmlRenderOptions CollapsedGeometryOptions() {
        HtmlRenderOptions options = TableGeometryOptions();
        options.Fonts.Add("Witness", File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
            "OfficeIMO.TestAssets", "Fonts", "OfficeIMOBaselineSans-Regular.ttf")));
        return options;
    }

    private static HtmlRenderDocument RenderCollapsedGeometry(string body) =>
        HtmlRenderTestDriver.Render(CollapsedGeometrySource(body), CollapsedGeometryOptions());

    private static string CollapsedGeometryNeighbors(string left, string right) =>
        "<table id='grid' style='width:200px;table-layout:fixed;border-collapse:collapse'><tr>"
        + "<td id='a' style='border-right:" + left + "'><div id='a-probe' class='probe'>Left</div></td>"
        + "<td id='b' style='border-left:" + right + "'><div id='b-probe' class='probe'>Right</div></td></tr></table>"
        + CollapsedGeometryFollowing;

    private static IEnumerable<HtmlRenderVisual> CollapsedGeometryPageVisuals(HtmlRenderPage page) => EnumerateRenderVisuals(page.Scene);

    private static void AssertCollapsedGeometryEquivalent(HtmlRenderDocument expected, HtmlRenderDocument actual, params string[] sources) {
        Assert.Equal(expected.Pages.Count, actual.Pages.Count);
        foreach (string source in sources) {
            HtmlRenderShape first = TableGeometryShape(expected, source);
            HtmlRenderShape second = TableGeometryShape(actual, source);
            Assert.Equal(first.X, second.X, 3);
            Assert.Equal(first.Y, second.Y, 3);
            Assert.Equal(first.Width, second.Width, 3);
            Assert.Equal(first.Height, second.Height, 3);
        }
        HtmlRenderText[] expectedText = expected.Pages.SelectMany(CollapsedGeometryPageVisuals).OfType<HtmlRenderText>().ToArray();
        HtmlRenderText[] actualText = actual.Pages.SelectMany(CollapsedGeometryPageVisuals).OfType<HtmlRenderText>().ToArray();
        Assert.Equal(expectedText.Length, actualText.Length);
        for (int index = 0; index < expectedText.Length; index++) {
            Assert.Equal(expectedText[index].Text, actualText[index].Text);
            Assert.Equal(expectedText[index].X, actualText[index].X, 3);
            Assert.Equal(expectedText[index].Y, actualText[index].Y, 3);
            Assert.Equal(expectedText[index].Width, actualText[index].Width, 3);
        }
        expected.RequireNoLoss();
        actual.RequireNoLoss();
    }
}

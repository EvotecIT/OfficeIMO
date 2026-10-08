using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlTables_ColspanContributionsDoNotDependOnSourceRowOrder() {
        const string spanning = "<tr><td colspan='2' style='width:400px'></td></tr>";
        const string singleColumns = "<tr><td id='left' style='width:300px'></td><td id='right' style='width:10px'></td></tr>";
        HtmlRenderDocument first = RenderTableIntrinsic("<table id='grid'>" + spanning + singleColumns + "</table>");
        HtmlRenderDocument last = RenderTableIntrinsic("<table id='grid'>" + singleColumns + spanning + "</table>");

        Assert.Equal(400D, TableIntrinsicGroup(first, "table#grid").Width, 3);
        Assert.Equal(TableIntrinsicGroup(first, "table#grid").Width, TableIntrinsicGroup(last, "table#grid").Width, 3);
        foreach (string source in new[] { "td#left", "td#right" }) {
            Assert.Equal(TableIntrinsicGroup(first, source).Width, TableIntrinsicGroup(last, source).Width, 3);
        }
    }

    [Fact]
    public void HtmlTables_OverlappingEqualColspansUseTheSameColumnBasis() {
        const string leftSpan = "<tr><td colspan='2' style='width:200px'></td><td style='width:20px'></td></tr>";
        const string rightSpan = "<tr><td style='width:40px'></td><td colspan='2' style='width:280px'></td></tr>";
        const string singles = "<tr><td id='left' style='width:80px'></td><td id='middle' style='width:60px'></td><td id='right' style='width:20px'></td></tr>";
        HtmlRenderDocument first = RenderTableIntrinsic("<table>" + leftSpan + rightSpan + singles + "</table>");
        HtmlRenderDocument reversed = RenderTableIntrinsic("<table>" + singles + rightSpan + leftSpan + "</table>");
        double[] columns = new[] { "td#left", "td#middle", "td#right" }
            .Select(source => TableIntrinsicGroup(first, source).Width).ToArray();

        Assert.True(columns[0] + columns[1] >= 200D);
        Assert.True(columns[1] + columns[2] >= 280D);
        foreach (string source in new[] { "td#left", "td#middle", "td#right" }) {
            Assert.Equal(TableIntrinsicGroup(first, source).Width, TableIntrinsicGroup(reversed, source).Width, 3);
        }
    }

    [Theory]
    [InlineData("auto")]
    [InlineData("fixed")]
    public void HtmlTables_ColSpanRepeatsItsDeclaredColumnWidth(string layout) {
        HtmlRenderDocument rendered = RenderTableIntrinsic("<table style='width:300px;table-layout:" + layout
            + "'><colgroup><col span='2' style='width:100px'><col></colgroup><tr>"
            + "<td id='left'></td><td id='middle'></td><td id='right'></td></tr></table>");

        Assert.Equal(100D, TableIntrinsicGroup(rendered, "td#left").Width, 3);
        Assert.Equal(100D, TableIntrinsicGroup(rendered, "td#middle").Width, 3);
        Assert.Equal(100D, TableIntrinsicGroup(rendered, "td#right").Width, 3);
    }

    [Theory]
    [InlineData(0D)]
    [InlineData(4D)]
    public void HtmlTables_ColspanWidthIncludesOnlyItsInternalSpacing(double spacing) {
        string value = spacing.ToString(System.Globalization.CultureInfo.InvariantCulture);
        HtmlRenderDocument rendered = RenderTableIntrinsic("<table id='grid' style='border-spacing:" + value
            + "px 0'><tr><td id='wide' colspan='2' style='width:100px'></td></tr>"
            + "<tr><td style='width:30px'></td><td style='width:30px'></td></tr></table>");

        Assert.Equal(100D, TableIntrinsicGroup(rendered, "td#wide").Width, 3);
        Assert.Equal(100D + spacing * 2D, TableIntrinsicGroup(rendered, "table#grid").Width, 3);
    }

    [Theory]
    [InlineData("styled")]
    [InlineData("blocks")]
    [InlineData("atomic")]
    [InlineData("generated")]
    public void HtmlTables_CellContributionsUseStyledContentAndLineBoundaries(string content) {
        var options = TableIntrinsicOptions();
        string body = content switch {
            "styled" => "<td id='cell' style='font-size:10px'><span style='font-size:40px'>WWWW</span></td>",
            "blocks" => "<td id='cell'><div>AAAAA</div><div>AAAAA</div></td>",
            "atomic" => "<td id='cell'>" + IntrinsicWidthAtoms + "</td>",
            _ => "<td id='cell'><span id='generated'>A</span></td>"
        };
        string html = "<style>#generated::before{content:'WWWW'}</style><table><tr>" + body + "</tr></table>";
        double expected = 90D;
        if (content != "atomic") {
            string text = content == "styled" || content == "generated" ? "WWWW" : "AAAAA";
            Assert.True(options.Fonts.TryMeasureText(text, content == "styled" ? 40D : 20D,
                "Pinned", OfficeFontStyle.Regular, out expected));
            if (content == "generated") {
                // Generated and authored text paint as distinct runs, so their
                // contribution must not borrow kerning across that boundary.
                Assert.True(options.Fonts.TryMeasureText("A", 20D, "Pinned", OfficeFontStyle.Regular, out double authored));
                expected += authored;
            }
        }

        HtmlRenderDocument rendered = RenderTableIntrinsic(html, options);

        Assert.Equal(expected, TableIntrinsicGroup(rendered, "td#cell").Width, 3);
        rendered.RequireNoLoss();
        if (content == "blocks") {
            HtmlRenderText[] lines = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().Where(text => text.Text == "AAAAA").ToArray();
            Assert.Equal(2, lines.Length);
            Assert.True(lines[1].Y > lines[0].Y);
        }
    }

    [Fact]
    public void HtmlTables_PageHeightContinuationRetainsOriginalColumnContributions() {
        const string html = "<style>@page{size:300px 60px;margin:0}@page:left{size:300px 80px}td{height:30px}</style>"
            + "<table id='grid' style='width:300px'><tr><td id='a1' style='width:200px'>First</td><td>B</td></tr>"
            + "<tr><td id='a2'>Next</td><td>B</td></tr><tr><td id='a3'>Last</td><td>B</td></tr>"
            + "<tr><td id='a4'>End</td><td>B</td></tr></table><p style='margin:0'>AfterTable</p>";
        var options = TableIntrinsicOptions();
        options.Mode = HtmlRenderMode.Paged;
        HtmlRenderDocument rendered = RenderTableIntrinsic(html, options);

        Assert.True(rendered.Pages.Count > 1);
        foreach (string source in new[] { "td#a1", "td#a2", "td#a3", "td#a4" }) {
            Assert.Equal(200D, TableIntrinsicGroup(rendered, source).Width, 3);
        }
        foreach (string marker in new[] { "First", "Next", "Last", "End", "AfterTable" }) {
            Assert.Single(rendered.Pages.SelectMany(page => page.Visuals.OfType<HtmlRenderText>()), text => text.Text == marker);
        }
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PagePseudoGeometryPending);
        HtmlRenderPage after = Assert.Single(rendered.Pages,
            page => page.Visuals.OfType<HtmlRenderText>().Any(text => text.Text == "AfterTable"));
        HtmlRenderPage lastRow = Assert.Single(rendered.Pages,
            page => page.Visuals.OfType<HtmlRenderText>().Any(text => text.Text == "End"));
        Assert.True(after.PageNumber >= lastRow.PageNumber);
        if (after.PageNumber == lastRow.PageNumber) {
            HtmlRenderText following = Assert.Single(after.Visuals.OfType<HtmlRenderText>(), text => text.Text == "AfterTable");
            HtmlRenderSemanticGroup cell = TableIntrinsicGroup(rendered, "td#a4");
            Assert.True(following.Y >= cell.Y + cell.Height - 0.001D);
        }
    }

    private static HtmlRenderDocument RenderTableIntrinsic(string body, HtmlRenderOptions? options = null) =>
        HtmlRenderTestDriver.Render("<!doctype html><style>body{margin:0;font:20px/24px Pinned}table{margin:0;border-spacing:0}"
            + "td{padding:0}div{margin:0}</style>" + body, options ?? TableIntrinsicOptions());

    private static HtmlRenderOptions TableIntrinsicOptions() {
        var options = new HtmlRenderOptions {
            ViewportWidth = 640D, Margins = HtmlRenderMargins.All(0D),
            UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser
        };
        options.Fonts.Add("Pinned", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "SourceSerif4-Regular.otf")));
        return options;
    }

    private static HtmlRenderSemanticGroup TableIntrinsicGroup(HtmlRenderDocument rendered, string source) =>
        Assert.Single(rendered.Pages.SelectMany(page => EnumerateTablePaginationScene(page.Scene))
            .OfType<HtmlRenderSemanticGroup>(), group => group.Source == source
                && group.Role is HtmlRenderSemanticGroupRole.Table or HtmlRenderSemanticGroupRole.TableCell);
}

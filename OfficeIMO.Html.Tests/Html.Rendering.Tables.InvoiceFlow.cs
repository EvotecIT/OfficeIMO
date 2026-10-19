using AngleSharp.Html.Parser;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void HtmlTables_PartialPageUsesRowBreaksUnlessWholeTableAvoidanceIsAuthored(bool avoid, bool nested) {
        string html = "<style>@page{size:180px 100px;margin:0}body{margin:0;font:8px/20px Arial}"
            + "table{width:160px;margin:0;border-collapse:collapse}th,td{padding:0;height:20px}</style>"
            + (nested ? "<main>" : string.Empty)
            + "<div style='height:40px'>Lead</div><table style='" + (avoid ? "break-inside:avoid" : string.Empty) + "'>"
            + "<thead><tr><th>Header</th></tr></thead><tbody>"
            + "<tr><td>Row0</td></tr><tr><td>Row1</td></tr><tr><td>Row2</td></tr></tbody>"
            + "<tfoot><tr><td>Footer</td></tr></tfoot></table>"
            + (nested ? "</main>" : string.Empty);

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        Assert.Equal(2, rendered.Pages.Count);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Lead");
        Assert.Equal(!avoid, rendered.Pages[0].Visuals.OfType<HtmlRenderText>().Any(text => text.Text == "Row0"));
        foreach (string marker in new[] { "Row0", "Row1", "Row2" }) {
            Assert.Single(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), text => text.Text == marker);
        }
        HtmlRenderPage[] bodyPages = rendered.Pages.Where(page => page.Visuals.OfType<HtmlRenderText>()
            .Any(text => text.Text.StartsWith("Row", StringComparison.Ordinal))).ToArray();
        Assert.Equal(avoid ? 1 : 2, bodyPages.Length);
        Assert.All(bodyPages, page => {
            Assert.Contains(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "Header");
            Assert.Contains(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "Footer");
        });
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment);
    }

    [Fact]
    public void HtmlTables_AutoWidthTotalsUseTheirMinimumAndFlexEndPlacementWithoutInventingBorders() {
        const string html = "<style>body{margin:0;font:10px Arial}.summary{display:flex;justify-content:flex-end;width:400px}"
            + "table{margin:0;min-width:260px;border-collapse:collapse;background:#eee}td{padding:4px 8px}"
            + ".number{text-align:right}.total td{border-top:2px solid blue}</style><div class='summary'>"
            + "<table id='totals'><tr><td>Net total</td><td class='number'>69,076.50 EUR</td></tr>"
            + "<tr class='total'><td>Total due</td><td class='number'>83,172.15 EUR</td></tr></table></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, InvoiceTableOptions(400D));
        HtmlRenderShape table = FindFlexShape(rendered, "table#totals");

        Assert.Equal(260D, table.Width, 3);
        Assert.Equal(140D, table.X, 3);
        HtmlRenderShape[] borders = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .Where(shape => shape.Shape.StrokeWidth > 0D).ToArray();
        Assert.NotEmpty(borders);
        Assert.All(borders, border => {
            Assert.StartsWith("table#totals:collapsed-border-h-1-", border.Source);
            Assert.Equal(OfficeColor.Blue, border.Shape.StrokeColor);
        });
    }

    [Fact]
    public void HtmlTables_FlexAutomaticMinimumIncludesEverySourceColumn() {
        const string html = "<style>body{margin:0;font:10px Arial}table{margin:0;border-collapse:collapse;background:#eee}"
            + "td{width:90px;padding:0}</style><div style='display:flex;width:200px'>"
            + "<table id='columns'><tr><td>A</td><td>B</td></tr></table>"
            + "<div id='neighbor' style='flex:0 0 80px;height:20px;background:blue'></div></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, InvoiceTableOptions(300D));
        HtmlRenderShape table = FindFlexShape(rendered, "table#columns");
        HtmlRenderShape neighbor = FindFlexShape(rendered, "div#neighbor");

        Assert.Equal(180D, table.Width, 3);
        Assert.Equal(table.X + table.Width, neighbor.X, 3);
        Assert.Equal(80D, neighbor.Width, 3);
    }

    [Fact]
    public void HtmlTables_FlexSpecifiedWidthCanShrinkToTheContentMinimum() {
        const string html = "<style>body{margin:0;font:10px Arial}table{margin:0;border-collapse:collapse;background:#eee}"
            + "td{padding:0}</style><div style='display:flex;width:300px'>"
            + "<table id='shrink' style='width:300px'><tr><td>Short</td></tr></table>"
            + "<div id='neighbor' style='flex:0 0 100px;height:20px;background:blue'></div></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, InvoiceTableOptions(300D));
        HtmlRenderShape table = FindFlexShape(rendered, "table#shrink");
        HtmlRenderShape neighbor = FindFlexShape(rendered, "div#neighbor");

        Assert.Equal(200D, table.Width, 3);
        Assert.Equal(table.X + table.Width, neighbor.X, 3);
    }

    [Theory]
    [InlineData("", "", "center", OfficeTextAlignment.Center)]
    [InlineData("start", "", "center", OfficeTextAlignment.Center)]
    [InlineData("initial", "", "center", OfficeTextAlignment.Center)]
    [InlineData("left", "", "left", OfficeTextAlignment.Left)]
    [InlineData("right", "", "right", OfficeTextAlignment.Right)]
    [InlineData("", "initial", "", OfficeTextAlignment.Left)]
    [InlineData("", "inherit", "", OfficeTextAlignment.Left)]
    [InlineData("", "unset", "", OfficeTextAlignment.Left)]
    [InlineData("", "revert", "center", OfficeTextAlignment.Center)]
    [InlineData("right", "inherit", "right", OfficeTextAlignment.Right)]
    [InlineData("right", "initial", "", OfficeTextAlignment.Left)]
    [InlineData("right", "revert", "right", OfficeTextAlignment.Right)]
    [InlineData("", "left", "left", OfficeTextAlignment.Left)]
    public void HtmlTables_HeaderAlignmentDefaultsRespectParentAlignmentAndAuthoredKeywords(
        string parentAlignment, string cellAlignment, string expectedComputed, OfficeTextAlignment expectedAlignment) {
        string html = "<style>body{margin:0;font:10px Arial}table{margin:0;width:160px;border-collapse:collapse}"
            + "th{padding:0}</style><table><tr style='"
            + (parentAlignment.Length == 0 ? string.Empty : "text-align:" + parentAlignment)
            + "'><th id='header' style='" + (cellAlignment.Length == 0 ? string.Empty : "text-align:" + cellAlignment)
            + "'><span id='nested'>Header</span></th></tr></table>";
        var source = new HtmlParser().ParseDocument(html);
        var styles = HtmlComputedStyleEngine.Compute(source);
        Assert.Equal(expectedComputed, styles[source.QuerySelector("#header")!].GetValue("text-align"));
        Assert.Equal(expectedComputed, styles[source.QuerySelector("#nested")!].GetValue("text-align"));

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, InvoiceTableOptions(180D));
        HtmlRenderText text = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>());
        HtmlRenderSemanticGroup cell = Assert.Single(EnumerateTablePaginationScene(rendered.Pages[0].Scene)
            .OfType<HtmlRenderSemanticGroup>(), group => group.Source == "th#header"
                && group.Role == HtmlRenderSemanticGroupRole.TableHeaderCell);
        double advance = Assert.IsType<double>(text.TextAdvanceWidth);
        double expectedX = cell.X + (expectedAlignment switch {
            OfficeTextAlignment.Center => (cell.Width - advance) / 2D,
            OfficeTextAlignment.Right => cell.Width - advance,
            _ => 0D
        });
        Assert.Equal(expectedX, text.X, 3);
        Assert.Equal(700, text.FontDescriptor.Weight);
    }

    private static HtmlRenderOptions InvoiceTableOptions(double width) => new HtmlRenderOptions {
        Mode = HtmlRenderMode.Continuous,
        ViewportWidth = width,
        ViewportHeight = 200D,
        Margins = HtmlRenderMargins.All(0D),
        AllowSystemFontFallback = false
    };
}

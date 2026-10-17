using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("width")]
    [InlineData("minimum")]
    [InlineData("caption")]
    [InlineData("flex")]
    public void HtmlTables_AutoSurplusPreservesNativeContentProportions(string sizing) {
        string css = sizing is "minimum" or "flex" ? "width:auto;min-width:600px"
            : sizing == "caption" ? "width:auto" : "width:600px";
        string caption = sizing == "caption" ? "<caption style='width:600px'>Caption</caption>" : "";
        string table = "<table style='" + css + "'>" + caption
            + "<tr><td><strong>Hgyp sample 0</strong></td><td>123</td></tr></table>";
        if (sizing == "flex") table = "<div style='display:flex;width:680px;justify-content:flex-end'>" + table + "</div>";

        HtmlRenderSemanticGroup grid = AutoDistributionTable(table);
        HtmlRenderSemanticGroup[] cells = AutoDistributionCells(grid);

        Assert.Equal(600D, grid.Width, 6);
        Assert.Equal(grid.Width, cells.Sum(cell => cell.Width), 6);
        // Edge155 with the same checked-in regular/bold font bytes assigns
        // 450.078125px to the first column. Allow subpixel layout rounding.
        Assert.InRange(cells[0].Width, 450.05D, 450.10D);
        Assert.InRange(cells[1].Width, 149.90D, 149.95D);
        Assert.Equal(cells[0].X + cells[0].Width, cells[1].X, 6);
    }

    [Theory]
    [InlineData("cell", 150D, 450D)]
    [InlineData("column", 150D, 450D)]
    [InlineData("percent-and-length", 150D, 450D)]
    [InlineData("two-percentages", 200D, 400D)]
    [InlineData("two-lengths", 207.407407407407D, 392.592592592593D)]
    public void HtmlTables_AutoSurplusKeepsPercentageAndLengthPriorities(string constraints, double first, double second) {
        string columns = constraints == "column" ? "<colgroup><col style='width:25%'><col></colgroup>" : "";
        string left = constraints switch {
            "column" => "",
            "two-lengths" => "width:100px",
            _ => "width:25%"
        };
        string right = constraints switch {
            "percent-and-length" => "width:100px",
            "two-percentages" => "width:50%",
            "two-lengths" => "width:200px",
            _ => ""
        };
        HtmlRenderSemanticGroup grid = AutoDistributionTable("<table style='width:600px'>" + columns
            + "<tr><td style='" + left + "'>A</td><td style='" + right + "'>B</td></tr></table>");
        HtmlRenderSemanticGroup[] cells = AutoDistributionCells(grid);

        Assert.Equal(first, cells[0].Width, 6);
        Assert.Equal(second, cells[1].Width, 6);
        Assert.Equal(600D, cells.Sum(cell => cell.Width), 6);
    }

    [Fact]
    public void HtmlTables_AutoColumnSpanRepeatsItsPercentageShare() {
        HtmlRenderSemanticGroup grid = AutoDistributionTable("<table style='width:600px'>"
            + "<colgroup><col span='2' style='width:25%'><col></colgroup>"
            + "<tr><td>A</td><td>B</td><td>123</td></tr></table>");

        Assert.Equal(new[] { 150D, 150D, 300D }, AutoDistributionCells(grid).Select(cell => cell.Width).ToArray());
    }

    [Theory]
    [InlineData("cells", 450D, 150D)]
    [InlineData("columns", 450D, 150D)]
    [InlineData("mixed-percentages", 450D, 150D)]
    [InlineData("length", 450D, 150D)]
    [InlineData("larger-than-table", 579.296875D, 20.703125D)]
    public void HtmlTables_ConflictingPercentagesPreserveTheTableAndContentMinimums(string constraints, double first, double second) {
        string columns = constraints == "columns" ? "<colgroup><col span='2' style='width:75%'></colgroup>" : "";
        string left = constraints == "columns" ? "" : constraints == "larger-than-table" ? "width:120%" : "width:75%";
        string right = constraints switch {
            "cells" => "width:75%",
            "mixed-percentages" => "width:50%",
            "length" => "width:200px",
            _ => ""
        };
        HtmlRenderSemanticGroup grid = AutoDistributionTable("<table style='width:600px'>" + columns
            + "<tr><td style='" + left + "'>A</td><td style='" + right + "'>B</td></tr></table>");
        HtmlRenderSemanticGroup[] cells = AutoDistributionCells(grid);

        Assert.Equal(600D, grid.Width, 6);
        Assert.Equal(first, cells[0].Width, 6);
        Assert.Equal(second, cells[1].Width, 6);
        Assert.Equal(grid.Width, cells.Sum(cell => cell.Width), 6);
    }

    [Theory]
    [InlineData("100px", 100D, 500D)]
    [InlineData("25%", 150D, 450D)]
    public void HtmlTables_EmptyAutomaticColumnRetainsItsSurplusPriority(string width, double first, double second) {
        HtmlRenderSemanticGroup grid = AutoDistributionTable("<table style='width:600px'>"
            + "<tr><td style='padding:0;width:" + width + "'>A</td><td style='padding:0'></td></tr></table>");
        HtmlRenderSemanticGroup[] cells = AutoDistributionCells(grid);

        Assert.Equal(first, cells[0].Width, 6);
        Assert.Equal(second, cells[1].Width, 6);
    }

    [Theory]
    [InlineData("", "123", 0D, 600D)]
    [InlineData("123", "", 600D, 0D)]
    [InlineData("", "", 300D, 300D)]
    public void HtmlTables_EmptyAutoColumnsDoNotInventContentWidth(string left, string right, double first, double second) {
        HtmlRenderSemanticGroup grid = AutoDistributionTable("<table style='width:600px'>"
            + "<tr><td style='padding:0'>" + left + "</td><td style='padding:0'>" + right + "</td></tr></table>");
        HtmlRenderSemanticGroup[] cells = AutoDistributionCells(grid);

        Assert.Equal(first, cells[0].Width, 6);
        Assert.Equal(second, cells[1].Width, 6);
        Assert.Equal(600D, cells.Sum(cell => cell.Width), 6);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlTables_AutoMinimumPreservesUnbreakableCellContributions(bool narrowContainer) {
        string table = "<table style='width:" + (narrowContainer ? "auto" : "40px")
            + "'><tr><td>UnbreakableIdentifier</td><td>123</td></tr></table>";
        if (narrowContainer) table = "<main style='width:40px'>" + table + "</main>";
        HtmlRenderSemanticGroup grid = AutoDistributionTable(table);
        HtmlRenderSemanticGroup[] cells = AutoDistributionCells(grid);

        Assert.Equal(155.1875D, cells[0].Width, 6);
        Assert.Equal(36.328125D, cells[1].Width, 6);
        Assert.Equal(191.515625D, grid.Width, 6);
    }

    private static HtmlRenderSemanticGroup AutoDistributionTable(string body) {
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 800D,
            Margins = HtmlRenderMargins.All(0D),
            DefaultFontFamily = "AutoWitness",
            AllowSystemFontFallback = false,
            UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser
        };
        options.Fonts.Add("AutoWitness", File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
            "OfficeIMO.TestAssets", "Fonts", "OfficeIMOBaselineSans-Regular.ttf")));
        options.Fonts.Add("AutoWitness", File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
            "OfficeIMO.TestAssets", "Fonts", "OfficeIMOBaselineSans-Bold.ttf")), OfficeFontStyle.Bold);
        HtmlRenderDocument document = HtmlRenderTestDriver.Render("<!doctype html><style>"
            + "body{margin:0;font:16px/20px AutoWitness;font-feature-settings:'kern' 0,'liga' 0}"
            + "table{margin:0;border-collapse:collapse}td{padding:5px 6px}"
            + "</style>" + body, options);
        return Assert.Single(document.Pages.SelectMany(page => EnumerateTablePaginationScene(page.Scene))
            .OfType<HtmlRenderSemanticGroup>(), item => item.Role == HtmlRenderSemanticGroupRole.Table);
    }

    private static HtmlRenderSemanticGroup[] AutoDistributionCells(HtmlRenderSemanticGroup table) =>
        EnumerateTablePaginationScene(table.Visuals).OfType<HtmlRenderSemanticGroup>()
            .Where(item => item.Role == HtmlRenderSemanticGroupRole.TableCell).OrderBy(item => item.X).ToArray();
}

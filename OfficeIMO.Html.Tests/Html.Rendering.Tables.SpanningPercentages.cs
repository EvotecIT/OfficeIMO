using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("50%", 150D, 150D, 300D)]
    [InlineData("75%", 225D, 225D, 150D)]
    [InlineData("100%", 300D, 300D, 0D)]
    [InlineData("120%", 360D, 240D, 0D)]
    [InlineData("calc(3 * 25%)", 225D, 225D, 150D)]
    [InlineData("attribute", 225D, 225D, 150D)]
    public void HtmlTables_SpanningPercentagesReserveTheirShareWithoutExpandingTheTable(
        string percentage, double first, double second, double third) {
        string width = percentage == "attribute" ? "width='75%'" : "style='width:" + percentage + "'";
        AssertSpanningPercentageTracks("<tr><td></td><td></td><td></td></tr>"
            + "<tr><td colspan='2' " + width + "></td><td></td></tr>", first, second, third);
    }

    [Theory]
    [InlineData("cell", "75%", 150D, 300D, 150D)]
    [InlineData("column", "75%", 150D, 300D, 150D)]
    [InlineData("larger", "50%", 450D, 75D, 75D)]
    public void HtmlTables_SpanningPercentagesUseOnlyTheUnassignedPercentage(
        string declaration, string percentage, double first, double second, double third) {
        string columns = declaration == "column" ? "<colgroup><col style='width:25%'><col><col></colgroup>" : "";
        string cellWidth = declaration == "column" ? "" : declaration == "larger" ? "width:75%" : "width:25%";
        AssertSpanningPercentageTracks(columns + "<tr><td style='" + cellWidth + "'></td><td></td><td></td></tr>"
            + "<tr><td colspan='2' style='width:" + percentage + "'></td><td></td></tr>", first, second, third);
    }

    [Theory]
    [InlineData(false, 225D, 225D, 150D)]
    [InlineData(true, 300D, 225D, 75D)]
    public void HtmlTables_SpanningPercentageConflictsRetainSourceColumnPriority(
        bool outsideFirst, double first, double second, double third) {
        string ordinary = outsideFirst ? "<td style='width:50%'></td><td></td><td></td>"
            : "<td></td><td></td><td style='width:50%'></td>";
        string spanning = outsideFirst ? "<td style='width:50%'></td><td colspan='2' style='width:75%'></td>"
            : "<td colspan='2' style='width:75%'></td><td style='width:50%'></td>";
        AssertSpanningPercentageTracks("<tr>" + ordinary + "</tr><tr>" + spanning + "</tr>", first, second, third);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlTables_SpanningPercentagesWeightPreferredLengths(bool columnDeclarations) {
        string columns = columnDeclarations
            ? "<colgroup><col style='width:100px'><col style='width:400px'><col></colgroup>" : "";
        string cells = columnDeclarations ? "<td></td><td></td>"
            : "<td style='width:100px'></td><td style='width:400px'></td>";
        AssertSpanningPercentageTracks(columns + "<tr>" + cells + "<td></td></tr>"
            + "<tr><td colspan='2' style='width:75%'></td><td></td></tr>", 90D, 360D, 150D);
    }

    [Fact]
    public void HtmlTables_SpanningPercentageGuessesPreserveOutsideUnbreakableContent() {
        const string content = "<span style='display:inline-block;width:500px;height:10px'></span>";
        AssertSpanningPercentageTracks("<tr><td></td><td></td><td>" + content + "</td></tr>"
            + "<tr><td colspan='2' style='width:75%'></td><td>" + content + "</td></tr>", 50D, 50D, 500D);
    }

    [Fact]
    public void HtmlTables_SpanningPercentagesPreserveAuthoredMinimumWidths() {
        AssertSpanningPercentageTracks("<tr><td style='min-width:100px'></td><td style='min-width:400px'></td><td></td></tr>"
            + "<tr><td colspan='2' style='width:75%'></td><td></td></tr>", 100D, 400D, 100D);
    }

    [Fact]
    public void HtmlTables_SpanningPercentageUsesTheAssignableWidthAndInternalSpacingOnce() {
        HtmlRenderSemanticGroup table = AutoDistributionTable("<style>"
            + "table{border-collapse:separate;border-spacing:4px}td{padding:0;height:20px}</style>"
            + "<table style='width:600px'><tr><td></td><td></td><td></td></tr>"
            + "<tr><td colspan='2' style='width:75%'></td><td></td></tr></table>");
        HtmlRenderSemanticGroup[] cells = AutoDistributionCells(table);
        HtmlRenderSemanticGroup[] tracks = cells.Where(cell => cell.ColumnSpan == 1)
            .GroupBy(cell => cell.Y).First(row => row.Count() == 3).OrderBy(cell => cell.X).ToArray();

        Assert.Equal(600D, table.Width, 6);
        Assert.Equal(219D, tracks[0].Width, 6);
        Assert.Equal(219D, tracks[1].Width, 6);
        Assert.Equal(146D, tracks[2].Width, 6);
        Assert.Equal(442D, Assert.Single(cells, cell => cell.ColumnSpan == 2).Width, 6);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlTables_EqualLengthSpanningPercentagesDoNotDependOnRowOrder(bool reverse) {
        const string first = "<tr><td colspan='2' style='width:75%'></td><td></td></tr>";
        const string second = "<tr><td></td><td colspan='2' style='width:75%'></td></tr>";
        AssertSpanningPercentageTracks("<tr><td></td><td></td><td></td></tr>"
            + (reverse ? second + first : first + second), 225D, 225D, 150D);
    }

    private static void AssertSpanningPercentageTracks(string rows, params double[] expected) {
        HtmlRenderSemanticGroup table = AutoDistributionTable("<style>td{padding:0;height:20px}</style>"
            + "<table style='width:600px'>" + rows + "</table>");
        HtmlRenderSemanticGroup[] tracks = AutoDistributionCells(table).Where(cell => cell.ColumnSpan == 1)
            .GroupBy(cell => cell.Y).First(row => row.Count() == expected.Length).OrderBy(cell => cell.X).ToArray();

        Assert.Equal(600D, table.Width, 6);
        Assert.Equal(table.Width, tracks.Sum(cell => cell.Width), 6);
        for (int index = 0; index < expected.Length; index++) {
            Assert.Equal(expected[index], tracks[index].Width, 6);
            if (index > 0) Assert.Equal(tracks[index - 1].X + tracks[index - 1].Width, tracks[index].X, 6);
        }
    }
}

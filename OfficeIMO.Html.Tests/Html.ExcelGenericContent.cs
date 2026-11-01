using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlExcelGenericContentTests {
    [Theory]
    [InlineData("<td><p>a</p><p>b</p></td>", "a\nb")]
    [InlineData("<td>\n  <p>a</p>\n  <p>b</p>\n</td>", "a\nb")]
    [InlineData("<td><div>a</div><div>b</div></td>", "a\nb")]
    [InlineData("<td><ul><li>a</li><li>b</li></ul></td>", "a\nb")]
    [InlineData("<td>a<div>b</div>c</td>", "a\nb\nc")]
    [InlineData("<td><table><tr><td>a</td></tr><tr><td>b</td></tr></table></td>", "a\nb")]
    [InlineData("<td><p>only</p></td>", "only")]
    [InlineData("<td>in<b>line</b></td>", "inline")]
    public void TableCellKeepsBlockContentOnSeparateLines(string cell, string expected) {
        using ExcelDocument workbook = Import("<table><tr>" + cell + "<td>x</td></tr></table>");

        Assert.True(workbook.Sheets[0].TryGetCellText(1, 1, out string text));
        Assert.Equal(expected, text);
    }

    [Fact]
    public void FooterRowFollowsTheBodyWhateverItsSourceOrder() {
        using ExcelDocument workbook = Import(
            "<table><tfoot><tr><td>Foot</td></tr></tfoot><tbody><tr><td>Body</td></tr></tbody><thead><tr><th>Head</th></tr></thead></table>");

        Assert.Equal(new[] { "Head", "Body", "Foot" }, Enumerable.Range(1, 3).Select(row => {
            Assert.True(workbook.Sheets[0].TryGetCellText(row, 1, out string text));
            return text;
        }));
    }

    [Theory]
    [InlineData("<style>.h{display:none}</style><p>a<span class=h>SECRET</span>b</p>", "ab")]
    [InlineData("<p>ab</p><p hidden>SECRET</p>", "ab")]
    [InlineData("<style>.h{display:none}</style><p>ab</p><div class=h><p>SECRET</p></div>", "ab")]
    public void HiddenContentIsNotImportedAndNoTitleRowIsInvented(string html, string expected) {
        using ExcelDocument workbook = Import(html);
        ExcelSheet sheet = Assert.Single(workbook.Sheets);

        Assert.Equal(expected, Assert.Single(sheet.EnumerateCells()).Value);
    }

    [Fact]
    public void HiddenTableCellLeavesTheVisibleCellsInPlace() {
        using ExcelDocument workbook = Import("<table><tr><td hidden>SECRET</td><td>a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr></table>");
        ExcelSheet sheet = workbook.Sheets[0];

        Assert.True(sheet.TryGetCellText(1, 1, out string first));
        Assert.True(sheet.TryGetCellText(1, 2, out string second));
        Assert.Equal(("a", "b"), (first, second));
        Assert.DoesNotContain(sheet.EnumerateCells(), cell => Equals(cell.Value, "SECRET"));
    }

    private static ExcelDocument Import(string html) =>
        HtmlConversionDocument.Parse(html)
            .ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic })
            .RequireValue();
}

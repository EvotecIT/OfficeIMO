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

    private static ExcelDocument Import(string html) =>
        HtmlConversionDocument.Parse(html)
            .ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic })
            .RequireValue();
}

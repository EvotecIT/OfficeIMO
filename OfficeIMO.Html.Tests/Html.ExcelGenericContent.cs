using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlExcelGenericContentTests {
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

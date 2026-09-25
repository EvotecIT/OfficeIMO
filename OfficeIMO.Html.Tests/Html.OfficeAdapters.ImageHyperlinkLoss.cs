using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlOfficeAdapters {
    [Fact]
    public void ExcelHtmlReportsLinkedPictureTargetLoss() {
        const string html = "<p><a href='https://www.cdc.gov/'><img alt='CDC' src='data:image/png;base64,AQID'></a></p>";
        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = result.RequireValue();

        Assert.Equal(1, result.Images);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.ContentOmitted
            && diagnostic.Source == "https://www.cdc.gov/");
    }

    [Fact]
    public void ExcelSemanticHtmlReportsLinkedInventoryPictureTargetLoss() {
        string image = "data:image/png;base64," + Convert.ToBase64String(OnePixelPng);
        string html = "<main><section class='officeimo-sheet' data-officeimo-sheet='Images'>"
            + "<section class='officeimo-images'><ul><li>"
            + "<a href='https://www.cdc.gov/'><img alt='CDC' src='" + image + "'></a>"
            + "</li></ul></section></section></main>";
        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Semantic });
        using ExcelDocument workbook = result.RequireValue();

        Assert.Equal(1, result.Images);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.ContentOmitted
            && diagnostic.Source == "https://www.cdc.gov/");
    }
}

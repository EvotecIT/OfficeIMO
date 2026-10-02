using OfficeIMO.Html;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using Xunit;
namespace OfficeIMO.Tests;
public sealed class HtmlExcelTableCaptionTests {
    private const string RegulatoryTableCaption = "List of National Secondary Drinking Water Regulations";
    private const string RegulatoryTableLink = "https://example.org/standards";
    private const string RegulatoryTableHtml = "<main><h1>Drinking water</h1><table><caption><a href='"
        + RegulatoryTableLink + "'>" + RegulatoryTableCaption + "</a></caption><tr><th>Contaminant</th><th>Standard</th></tr>"
        + "<tr><td>Fluoride</td><td>2.0 mg/L</td></tr></table></main>";

    [Fact]
    public void ExcelCaptionFollowsMergedRowExtent() {
        const string html = "<table><caption>Caption</caption><tr><td rowspan='5'>A</td></tr><tr></tr><tr></tr><tr></tr><tr></tr></table>";
        using var document = HtmlConversionDocument.Parse(html).ToExcelDocument(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        var sheet = Assert.Single(document.Sheets);
        Assert.Equal(5, Assert.Single(sheet.GetMergedRanges()).EndRow);
        Assert.Equal("Caption", sheet.CellAt(7, 1).GetValue<string>());
    }

    [Fact]
    public void ExcelCaptionLinkHonorsNativeMetadataLimit() {
        var limits = HtmlImportLimits.CreateDefault();
        limits.MaxMetadataCharacters = 64;
        string html = "<table><caption><a href='https://example.test/" + new string('x', 100) + "'>Units</a></caption><tr><td>m</td></tr></table>";
        var result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic, Limits = limits });
        using var document = result.RequireValue();
        Assert.Empty(Assert.Single(document.Sheets).GetHyperlinks());
        Assert.Equal("Units", document.Sheets[0].CellAt(3, 1).GetValue<string>());
        Assert.Contains(result.Report.Diagnostics, item => item.Code == HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded);
    }

    [Fact]
    public void ExcelRetainsFullLinkedCaptionAfterNativeTableCells() {
        HtmlToExcelResult result = HtmlConversionDocument.Parse(RegulatoryTableHtml).ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = result.RequireValue();
        using var artifact = new MemoryStream();
        workbook.Save(artifact);
        using ExcelDocument reopened = ExcelDocument.Load(new MemoryStream(artifact.ToArray()));

        ExcelSheet table = Assert.Single(reopened.Sheets);
        Assert.Equal("Contaminant", table.CellAt(1, 1).GetValue<string>());
        Assert.Equal("Fluoride", table.CellAt(2, 1).GetValue<string>());
        Assert.Equal(RegulatoryTableCaption, table.CellAt(4, 1).GetValue<string>());
        Assert.Equal(RegulatoryTableLink, table.GetHyperlinks()["A4"].Target);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.ContentApproximated
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation
            && diagnostic.Detail?.Contains("originalLength=" + RegulatoryTableCaption.Length,
                StringComparison.Ordinal) == true);
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.ContentOmitted
            && diagnostic.Message.Contains("caption", StringComparison.OrdinalIgnoreCase));
    }
}

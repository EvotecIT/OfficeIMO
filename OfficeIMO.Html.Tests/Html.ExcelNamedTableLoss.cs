using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlExcelNamedTableLossTests {
    [Theory]
    [InlineData(false, ExcelHtmlExportProfile.SemanticTables)]
    [InlineData(true, ExcelHtmlExportProfile.SemanticTables)]
    [InlineData(false, ExcelHtmlExportProfile.VisualReview)]
    [InlineData(true, ExcelHtmlExportProfile.VisualReview)]
    public void NativeTableMetadataOmissionIsReportedPerExportScope(bool worksheetOnly, ExcelHtmlExportProfile profile) {
        using var workbook = ExcelDocument.Create();
        var orders = AddTableSheet(workbook, "Orders", "OrderItems");
        AddTableSheet(workbook, "Inventory", "StockItems");
        var options = new ExcelHtmlSaveOptions { ExportProfile = profile };
        var result = worksheetOnly ? orders.ToHtmlResult(options) : workbook.ToHtmlResult(options);
        var omitted = result.Report.Diagnostics.Where(d => d.Source?.StartsWith("excel:table:", StringComparison.Ordinal) == true).ToArray();

        Assert.True(result.Report.HasLoss);
        Assert.Equal(worksheetOnly ? 1 : 2, omitted.Length);
        Assert.All(omitted, d => {
            Assert.Equal(HtmlConversionDiagnosticCodes.ContentOmitted, d.Code);
            Assert.Equal(OfficeConversionLossKind.Omission, d.LossKind);
            Assert.Contains("range=A1:B2", d.Detail);
        });
        Assert.Contains(omitted, d => d.Source == "excel:table:Orders/OrderItems");
        if (worksheetOnly) Assert.DoesNotContain(omitted, d => d.Source == "excel:table:Inventory/StockItems");
    }

    [Fact]
    public void UnrelatedPlainWorksheetHasNoNamedTableOmission() {
        using var workbook = ExcelDocument.Create();
        AddTableSheet(workbook, "Orders", "OrderItems");
        var plain = workbook.AddWorksheet("Plain");
        plain.CellValue(1, 1, "Preserved plain value");
        var result = plain.ToHtmlResult();
        Assert.False(result.Report.HasLoss);
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Source?.StartsWith("excel:table:", StringComparison.Ordinal) == true);
        Assert.Contains("Preserved plain value", result.Value);
    }

    [Fact]
    public void SemanticNamedTableExportReportsLossWhileCellsSurviveNativeReopen() {
        using var source = ExcelDocument.Create();
        AddTableSheet(source, "Orders", "OrderItems");
        var exported = source.ToHtmlResult();
        Assert.True(exported.Report.HasLoss);
        var imported = HtmlConversionDocument.Parse(exported.Value, new HtmlConversionDocumentOptions { Trust = HtmlInputTrust.Trusted }).ToExcelDocumentResult();
        using var restored = imported.RequireValue();
        using var bytes = new MemoryStream();
        restored.Save(bytes);
        using var reopened = ExcelDocument.Load(new MemoryStream(bytes.ToArray()));
        Assert.Empty(reopened.GetTables());
        Assert.Equal("Item", reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
        Assert.Equal("Alpha", reopened.Sheets[0].CellAt(2, 1).GetValue<string>());
        Assert.Equal(2, reopened.Sheets[0].CellAt(2, 2).GetValue<int>());
    }

    private static ExcelSheet AddTableSheet(ExcelDocument workbook, string sheetName, string tableName) {
        var sheet = workbook.AddWorksheet(sheetName);
        sheet.CellValue(1, 1, "Item");
        sheet.CellValue(1, 2, "Count");
        sheet.CellValue(2, 1, "Alpha");
        sheet.CellValue(2, 2, 2);
        sheet.AddTable("A1:B2", true, tableName, ExcelTableStyle.TableStyleMedium9);
        return sheet;
    }
}

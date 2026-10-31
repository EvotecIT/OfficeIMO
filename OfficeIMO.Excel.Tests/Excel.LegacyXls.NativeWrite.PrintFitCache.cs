using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false)]
    [InlineData(null)]
    [InlineData(true)]
    public void LegacyXls_NativeSave_PreservesStoredFitDimensionsAndScalingMode(bool? fitEnabled) {
        string xlsx = Path.Combine(_directoryWithFiles, "StoredPrintFit" + fitEnabled + ".xlsx");
        string xls = Path.ChangeExtension(xlsx, ".xls");
        using (var workbook = ExcelDocument.Create(xlsx, new ExcelCreateOptions { FirstSheetName = "Print" })) {
            ExcelSheet sheet = workbook.Sheets[0];
            sheet.CellValue(1, 1, "Cached fit dimensions");
            sheet.SetPageSetup(fitToWidth: 2, fitToHeight: 0, scale: 85);
            PageSetupProperties properties = sheet.WorksheetPart.Worksheet
                .GetFirstChild<SheetProperties>()!.GetFirstChild<PageSetupProperties>()!;
            properties.FitToPage = fitEnabled;
            workbook.Save(xls);
        }
        using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xls);
        result.EnsureNoImportErrors();
        var stored = Assert.Single(result.Workbook.Worksheets).PageSetup!;
        Assert.Equal(fitEnabled == true, stored.FitToPage);
        Assert.Equal((ushort)85, stored.Scale);
        Assert.Equal((ushort)2, stored.FitToWidth);
        Assert.Equal((ushort)0, stored.FitToHeight);
        ExcelSheetPageSetup effective = result.Document.Sheets[0].GetPageSetup();
        Assert.Equal(fitEnabled == true ? 2U : (uint?)null, effective.FitToWidth);
        Assert.Equal(fitEnabled == true ? 0U : (uint?)null, effective.FitToHeight);
    }
}

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(0, 1U, 0U, 1)]
    [InlineData(1, 1U, 1U, 1)]
    [InlineData(2, null, null, 3)]
    [InlineData(3, null, null, 3)]
    [InlineData(4, 0U, 1U, 3)]
    public void SaveAsPdf_ResolvesNativePrintFitDefaultsAndEnabledState(int settingsCase, uint? fitWidth, uint? fitHeight, int expectedPages) {
        string path = Path.Combine(_directoryWithFiles, "NativePrintDefaults" + settingsCase + ".xlsx");
        using (var workbook = ExcelDocument.Create(path, new ExcelCreateOptions { FirstSheetName = "Layout" })) {
            ExcelSheet sheet = workbook.Sheets[0];
            sheet.SetColumnWidth(1, 50);
            sheet.SetColumnWidth(2, 80);
            sheet.SetColumnWidth(3, 40);
            for (int row = 1; row <= 24; row++) {
                for (int column = 1; column <= 3; column++) sheet.CellValue(row, column, "Cell");
                sheet.SetRowHeight(row, 24);
            }
            sheet.SetPageSetup(fitToWidth: 1, fitToHeight: 0, paperSize: ExcelPaperSize.A4);
            workbook.Save();
        }
        // Native Excel omits dimensions that equal the format defaults and
        // retains cached dimensions when manual print scaling is enabled.
        using (var package = SpreadsheetDocument.Open(path, true)) {
            Worksheet worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet;
            PageSetup setup = worksheet.GetFirstChild<PageSetup>()!;
            PageSetupProperties properties = worksheet.GetFirstChild<SheetProperties>()!.GetFirstChild<PageSetupProperties>()!;
            if (settingsCase == 2) properties.FitToPage = false;
            if (settingsCase == 3) properties.FitToPage = null;
            setup.Scale = 100;
            if (settingsCase <= 1) setup.FitToWidth = null;
            if (settingsCase == 1 || settingsCase == 4) setup.FitToHeight = null;
            if (settingsCase == 4) setup.FitToWidth = 0;
            worksheet.Save();
        }
        using var loaded = ExcelDocument.Load(path);
        ExcelSheetPageSetup pageSetup = loaded.Sheets[0].GetPageSetup();
        Assert.Equal(fitWidth, pageSetup.FitToWidth);
        Assert.Equal(fitHeight, pageSetup.FitToHeight);
        PdfCore.PdfDocumentInfo pdf = PdfCore.PdfInspector.Inspect(loaded.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(expectedPages, pdf.Pages.Count);
    }
}

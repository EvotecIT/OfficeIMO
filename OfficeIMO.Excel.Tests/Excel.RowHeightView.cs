using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Xlsb_DefaultViewRetainsGridlineVisibilityAcrossRewrite(bool visible) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        sheet.Cell(1, 1, "Original");
        sheet.SetGridlinesVisible(visible);
        using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(document.ToBytes(ExcelFileFormat.Xlsb), writable: false));
        Assert.Equal(visible, loaded.Sheets[0].GetViewInfo().ShowGridlines);
        loaded.Sheets[0].Cell(1, 1, "Edited");
        using ExcelDocument reopened = ExcelDocument.Load(new MemoryStream(loaded.ToBytes(ExcelFileFormat.Xlsb), writable: false));
        Assert.Equal(visible, reopened.Sheets[0].GetViewInfo().ShowGridlines);
        Assert.Equal("Edited", reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void RowHeight_SaveRetainsViewRequiredForExcelToReadPointHeights(bool validate, bool unfreeze) {
        using var stream = new MemoryStream();
        using (ExcelDocument document = ExcelDocument.Create()) {
            ExcelSheet sheet = document.AddWorksheet("Report");
            sheet.Cell(1, 1, "Explicit height");
            sheet.SetRowHeight(1, 24D);
            if (unfreeze) { sheet.Freeze(1, 1); sheet.Freeze(); }
            document.Save(stream, new ExcelSaveOptions { ValidateOpenXml = validate });
        }
        stream.Position = 0;
        using SpreadsheetDocument package = SpreadsheetDocument.Open(stream, false);
        Worksheet worksheet = Assert.Single(package.WorkbookPart!.WorksheetParts).Worksheet;
        SheetView view = Assert.Single(worksheet.GetFirstChild<SheetViews>()!.Elements<SheetView>());
        Assert.Equal(0U, view.WorkbookViewId!.Value);
        Assert.Null(view.GetFirstChild<Pane>());
        Assert.Equal(24D, Assert.Single(worksheet.Descendants<Row>()).Height!.Value);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(package));
    }
}

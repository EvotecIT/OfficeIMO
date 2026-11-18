using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    /// <summary>
    /// Excel opens a workbook on its active tab and unhides that tab when it is hidden, so hiding the active sheet
    /// has to move the activation to a visible sheet in every saved format.
    /// </summary>
    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    [InlineData(".xls")]
    public void HidingTheActiveWorksheet_ActivatesTheNextVisibleWorksheetInEveryFormat(string extension) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.HiddenActive.{Guid.NewGuid():N}{extension}");
        try {
            using (var document = ExcelDocument.Create(path)) {
                document.AddSheet("Lookup").Cell(1, 1).SetValue("hidden");
                document.AddSheet("Report").Cell(1, 1).SetValue("shown");
                document.Sheets[0].SetHidden(true);
                Assert.Equal(1, document.ActiveWorksheetIndex);
                document.Save();
            }

            if (extension == ".xlsb") {
                // BrtBookView stores the active tab (itabCur) after five 32-bit window fields and itabFirst.
                using var workbookStream = new MemoryStream(ReadZipEntry(File.ReadAllBytes(path), "xl/workbook.bin"));
                var bookView = Assert.Single(OfficeIMO.Excel.Xlsb.Biff12.XlsbRecordReader.ReadAll(workbookStream), record => record.Type == 158);
                Assert.Equal(1U, BitConverter.ToUInt32(bookView.Data, 24));
            }

            using var loaded = ExcelDocument.Load(path);
            Assert.True(loaded.Sheets[0].Hidden);
            Assert.Equal(1, loaded.ActiveWorksheetIndex);
            Assert.Equal(new bool?[] { false, true }, loaded.Sheets
                .Select(sheet => sheet.WorksheetPart.Worksheet.GetFirstChild<SheetViews>()?.Elements<SheetView>().FirstOrDefault()?.TabSelected?.Value)
                .ToArray());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void HidingWorksheets_MovesActivationOnlyOffTheActiveWorksheet() {
        using var document = ExcelDocument.Create();
        document.AddSheet("First").Cell(1, 1).SetValue(1);
        document.AddSheet("Second").Cell(1, 1).SetValue(2);
        document.AddSheet("Third").Cell(1, 1).SetValue(3);
        document.SetActiveWorksheet(2);

        document.Sheets[2].SetVeryHidden(true);
        Assert.Equal(1, document.ActiveWorksheetIndex);

        document.Sheets[0].SetHidden(true);
        Assert.Equal(1, document.ActiveWorksheetIndex);
    }

    [Fact]
    public void HidingTheActiveWorksheet_MovesEveryWorkbookViewOnIt() {
        using var document = ExcelDocument.Create();
        document.AddSheet("First").Cell(1, 1).SetValue(1);
        document.AddSheet("Second").Cell(1, 1).SetValue(2);
        document.AddSheet("Third").Cell(1, 1).SetValue(3);
        document.SetActiveWorksheet(0);
        BookViews views = document.WorkbookRoot.GetFirstChild<BookViews>()!;
        views.Append(new WorkbookView { ActiveTab = 0U }, new WorkbookView { ActiveTab = 2U });

        document.Sheets[0].SetHidden(true);

        Assert.Equal(new uint?[] { 1U, 1U, 2U }, views.Elements<WorkbookView>().Select(view => view.ActiveTab?.Value).ToArray());
    }

    [Fact]
    public void HidingTheOnlyVisibleWorksheet_LeavesTheActivationUnchanged() {
        using var document = ExcelDocument.Create();
        document.AddSheet("First").Cell(1, 1).SetValue(1);
        document.AddSheet("Second").Cell(1, 1).SetValue(2);
        document.Sheets[1].SetHidden(true);

        document.Sheets[0].SetHidden(true);

        Assert.Equal(0, document.ActiveWorksheetIndex);
    }
}

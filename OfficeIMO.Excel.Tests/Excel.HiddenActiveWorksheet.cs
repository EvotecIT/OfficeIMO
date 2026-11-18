using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    /// <summary>
    /// Excel opens a workbook on its active tab and unhides that tab when it is hidden, so hiding the active sheet
    /// has to move the activation to a visible sheet. Each case is saved to every format and read back.
    /// </summary>
    [Theory]
    [InlineData("plain", ".xlsx")]
    [InlineData("plain", ".xlsb")]
    [InlineData("plain", ".xls")]
    [InlineData("frozen", ".xlsx")]
    [InlineData("frozen", ".xlsb")]
    [InlineData("frozen", ".xls")]
    [InlineData("veryHidden", ".xlsx")]
    [InlineData("veryHidden", ".xlsb")]
    [InlineData("veryHidden", ".xls")]
    [InlineData("loadedXlsb", ".xlsx")]
    [InlineData("loadedXlsb", ".xlsb")]
    [InlineData("loadedXlsb", ".xls")]
    [InlineData("twoViews", ".xlsx")]
    [InlineData("twoViews", ".xlsb")]
    [InlineData("twoViews", ".xls")]
    public void HidingTheActiveWorksheet_ActivatesAVisibleSheetInEveryFormat(string scenario, string extension) {
        string directory = Path.Combine(Path.GetTempPath(), $"OfficeIMO.HiddenActive.{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try {
            string path = Path.Combine(directory, "result" + extension);
            ExcelDocument document;
            if (scenario == "loadedXlsb") {
                // A workbook read from XLSB and saved again goes through the XLSB preservation checks.
                string source = Path.Combine(directory, "source.xlsb");
                using (ExcelDocument created = CreateHiddenActiveWorkbook(source, freeze: true)) created.Save();
                document = ExcelDocument.Load(source);
            } else {
                document = CreateHiddenActiveWorkbook(null, freeze: scenario == "frozen");
            }

            using (document) {
                if (scenario == "twoViews") {
                    document.SetActiveWorksheet(0);
                    document.WorkbookRoot.GetFirstChild<BookViews>()!.Append(new WorkbookView { ActiveTab = 0U });
                }

                if (scenario == "veryHidden") document.Sheets[0].SetVeryHidden(true);
                else document.Sheets[0].SetHidden(true);
                Assert.Equal(1, document.ActiveWorksheetIndex);
                document.Save(path);
            }

            uint[] expectedViews = scenario == "twoViews" ? new[] { 1U, 1U } : new[] { 1U };
            if (extension == ".xlsb") {
                // BrtBookView stores the active tab (itabCur) after five 32-bit window fields and itabFirst.
                using var workbookStream = new MemoryStream(ReadZipEntry(File.ReadAllBytes(path), "xl/workbook.bin"));
                uint[] activeTabs = OfficeIMO.Excel.Xlsb.Biff12.XlsbRecordReader.ReadAll(workbookStream)
                    .Where(record => record.Type == 158).Select(record => BitConverter.ToUInt32(record.Data, 24)).ToArray();
                Assert.Equal(expectedViews, activeTabs);
            }

            if (extension == ".xlsx") {
                using SpreadsheetDocument package = SpreadsheetDocument.Open(path, false);
                Assert.Empty(new OpenXmlValidator().Validate(package));
                Assert.Equal(expectedViews.Select(tab => (uint?)tab), package.WorkbookPart!.Workbook.GetFirstChild<BookViews>()!
                    .Elements<WorkbookView>().Select(view => view.ActiveTab?.Value));
            }

            using var loaded = ExcelDocument.Load(path);
            Assert.True(loaded.Sheets[0].Hidden);
            Assert.Equal(scenario == "veryHidden", loaded.Sheets[0].VeryHidden);
            Assert.False(loaded.Sheets[1].Hidden);
            Assert.Equal(1, loaded.ActiveWorksheetIndex);
            if (scenario != "loadedXlsb" && extension != ".xlsb") {
                // The XLSB reader does not carry tab selection; the other formats keep it on the new active sheet.
                Assert.Equal(new[] { false, true }, loaded.Sheets.Select(sheet => sheet.WorksheetPart.Worksheet
                    .GetFirstChild<SheetViews>()?.Elements<SheetView>().FirstOrDefault()?.TabSelected?.Value == true).ToArray());
            }
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static ExcelDocument CreateHiddenActiveWorkbook(string? path, bool freeze) {
        ExcelDocument document = path == null ? ExcelDocument.Create() : ExcelDocument.Create(path);
        ExcelSheet lookup = document.AddSheet("Lookup");
        lookup.Cell(1, 1).SetValue("hidden");
        lookup.Cell(2, 1).SetValue(2);
        if (freeze) lookup.View.Freeze(topRows: 1);
        document.AddSheet("Report").Cell(1, 1).SetValue("shown");
        return document;
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    [InlineData(".xls")]
    public void HidingEverySheet_LeavesTheActivationAndSavesInEveryFormat(string extension) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.AllHidden.{Guid.NewGuid():N}{extension}");
        try {
            using (var document = ExcelDocument.Create()) {
                document.AddSheet("First").Cell(1, 1).SetValue(1);
                document.AddSheet("Second").Cell(1, 1).SetValue(2);
                document.Sheets[1].SetHidden(true);
                Assert.Equal(0, document.ActiveWorksheetIndex);

                document.Sheets[0].SetHidden(true);

                Assert.True(document.Sheets.All(sheet => sheet.Hidden));
                Assert.Equal(0, document.ActiveWorksheetIndex);
                document.Save(path);
            }

            using var loaded = ExcelDocument.Load(path);
            Assert.True(loaded.Sheets.All(sheet => sheet.Hidden));
            Assert.Equal(0, loaded.ActiveWorksheetIndex);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void HidingASheet_MovesOnlyViewsActiveOnThatSheet() {
        using var document = ExcelDocument.Create();
        document.AddSheet("First").Cell(1, 1).SetValue(1);
        document.AddSheet("Second").Cell(1, 1).SetValue(2);
        document.AddSheet("Third").Cell(1, 1).SetValue(3);
        document.SetActiveWorksheet(0);
        // First is hidden last, when no other sheet is visible, so it stays active.
        document.Sheets[1].SetHidden(true);
        document.Sheets[2].SetHidden(true);
        document.Sheets[0].SetHidden(true);
        Assert.Equal(0, document.ActiveWorksheetIndex);
        document.Sheets[1].SetHidden(false);
        document.Sheets[2].SetHidden(false);

        // The active sheet was already hidden; hiding another sheet leaves the activation where it is.
        document.Sheets[2].SetHidden(true);

        Assert.Equal(0, document.ActiveWorksheetIndex);
        Assert.True(document.Sheets[2].Hidden);
    }

    [Fact]
    public void HidingTheActiveWorksheet_MovesEachViewAndItsSelection() {
        using var document = ExcelDocument.Create();
        document.AddSheet("First").Cell(1, 1).SetValue(1);
        document.AddSheet("Second").Cell(1, 1).SetValue(2);
        document.AddSheet("Third").Cell(1, 1).SetValue(3);
        document.SetActiveWorksheet(2);
        BookViews views = document.WorkbookRoot.GetFirstChild<BookViews>()!;
        views.Append(new WorkbookView { ActiveTab = 0U });
        Worksheet first = document.Sheets[0].WorksheetPart.Worksheet;
        first.GetFirstChild<SheetViews>()!.Append(new SheetView { WorkbookViewId = 1U, TabSelected = true });
        document.Sheets[1].WorksheetPart.Worksheet.GetFirstChild<SheetViews>()!.Append(new SheetView { WorkbookViewId = 1U });

        document.Sheets[0].SetHidden(true);

        Assert.Equal(new uint?[] { 2U, 1U }, views.Elements<WorkbookView>().Select(view => view.ActiveTab?.Value).ToArray());
        SheetView? SelectionOf(int sheet, uint workbookView) => document.Sheets[sheet].WorksheetPart.Worksheet
            .GetFirstChild<SheetViews>()?.Elements<SheetView>().FirstOrDefault(view => view.WorkbookViewId?.Value == workbookView);
        Assert.NotEqual(true, SelectionOf(0, 1U)?.TabSelected?.Value);
        Assert.True(SelectionOf(1, 1U)?.TabSelected?.Value);
        Assert.True(SelectionOf(2, 0U)?.TabSelected?.Value);
        Assert.NotEqual(true, SelectionOf(1, 0U)?.TabSelected?.Value);
    }

    [Fact]
    public void HidingTheActiveWorksheet_SelectsAVisibleChartSheet() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.HiddenChartTab.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                document.AddSheet("Data").Cell(1, 1).SetValue(1);
                document.AddSheet("Other").Cell(1, 1).SetValue(2);
                document.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                WorkbookPart workbookPart = package.WorkbookPart!;
                ChartsheetPart chartsheetPart = workbookPart.AddNewPart<ChartsheetPart>();
                DrawingsPart drawings = chartsheetPart.AddNewPart<DrawingsPart>();
                drawings.WorksheetDrawing = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();
                chartsheetPart.Chartsheet = new Chartsheet(
                    new ChartSheetViews(new ChartSheetView { WorkbookViewId = 0U }),
                    new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = chartsheetPart.GetIdOfPart(drawings) });
                Sheets sheets = workbookPart.Workbook.Sheets!;
                sheets.InsertAt(new Sheet { Id = workbookPart.GetIdOfPart(chartsheetPart), SheetId = 9U, Name = "Chart" }, 1);
                workbookPart.Workbook.Save();
            }

            using (var document = ExcelDocument.Load(path)) {
                document.Sheets[0].SetHidden(true);
                document.Save();
            }

            using SpreadsheetDocument saved = SpreadsheetDocument.Open(path, false);
            Assert.Empty(new OpenXmlValidator().Validate(saved));
            Assert.Equal(1U, saved.WorkbookPart!.Workbook.GetFirstChild<BookViews>()!.Elements<WorkbookView>().First().ActiveTab?.Value);
            Assert.True(saved.WorkbookPart.ChartsheetParts.Single().Chartsheet!
                .GetFirstChild<ChartSheetViews>()!.Elements<ChartSheetView>().Single().TabSelected?.Value);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    [InlineData(".xls")]
    public void LoadedXlsbWithAHiddenFirstSheet_SavesInEveryFormat(string extension) {
        string directory = Path.Combine(Path.GetTempPath(), $"OfficeIMO.HiddenFirstXlsb.{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try {
            string source = Path.Combine(directory, "source.xlsb");
            using (ExcelDocument created = CreateHiddenActiveWorkbook(source, freeze: true)) {
                created.Sheets[0].SetHidden(true);
                created.Save();
            }

            string path = Path.Combine(directory, "result" + extension);
            using (var loaded = ExcelDocument.Load(source)) {
                Assert.Equal(1, loaded.ActiveWorksheetIndex);
                loaded.Sheets[1].Cell(2, 1).SetValue("edited");
                loaded.Save(path);
            }

            using var reloaded = ExcelDocument.Load(path);
            Assert.True(reloaded.Sheets[0].Hidden);
            Assert.Equal(1, reloaded.ActiveWorksheetIndex);
            Assert.True(reloaded.Sheets[1].TryGetCellText(2, 1, out string text));
            Assert.Equal("edited", text);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void ActivatingAWorksheetWithoutViews_InsertsTheViewInSchemaOrder() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.ViewOrder.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                document.AddSheet("First").Cell(1, 1).SetValue(1);
                ExcelSheet second = document.AddSheet("Second");
                second.Cell(1, 1).SetValue(2);
                second.GroupColumns(2, 3);
                second.WorksheetPart.Worksheet.GetFirstChild<SheetViews>()!.Remove();
                document.Sheets[0].SetHidden(true);
                document.Save();
            }

            using SpreadsheetDocument package = SpreadsheetDocument.Open(path, false);
            Assert.Empty(new OpenXmlValidator().Validate(package));
        } finally {
            File.Delete(path);
        }
    }
}

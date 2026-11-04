using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_PrintLayoutPreset_AppliesReportWorkflowSettings() {
            string filePath = Path.Combine(_directoryWithFiles, "PrintLayoutPreset.Report.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Report");
                sheet.Cell(1, 1).SetValue("Region");
                sheet.Cell(2, 1).SetValue("EU");

                sheet.PageSetup.ApplyLayout(new ExcelPrintLayoutOptions {
                    Preset = ExcelPrintLayoutPreset.Report,
                    PrintArea = "A1:D25",
                    PaperSize = ExcelPaperSize.A4,
                    RepeatFirstColumn = 1,
                    RepeatLastColumn = 1
                });
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                ExcelSheet sheet = document["Report"];
                ExcelSheetPageSetup setup = sheet.PageSetup;
                Assert.Equal(OfficePageOrientation.Landscape, setup.Orientation);
                Assert.Equal(1U, setup.FitToWidth);
                Assert.Equal(0U, setup.FitToHeight);
                Assert.Equal(ExcelPageOrder.DownThenOver, setup.PageOrder);
                Assert.Equal(ExcelPaperSize.A4, setup.PaperSize);
                Assert.Equal(9U, setup.PaperSizeCode);
                Assert.NotNull(setup.Margins);
                Assert.Equal(0.25D, setup.Margins!.Left.ToInches());
                Assert.Equal("$A$1:$D$25", sheet.PageSetup.PrintArea);

                ExcelPrintTitles titles = sheet.PageSetup.PrintTitles;
                Assert.Equal(1, titles.FirstRow);
                Assert.Equal(1, titles.LastRow);
                Assert.Equal(1, titles.FirstColumn);
                Assert.Equal(1, titles.LastColumn);
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                Worksheet worksheet = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet;
                Assert.True(worksheet.GetFirstChild<SheetProperties>()?.GetFirstChild<PageSetupProperties>()?.FitToPage?.Value);
            }
        }

        [Fact]
        public void Test_PrintLayoutPreset_ClearsStaleFitToPageSettings() {
            string filePath = Path.Combine(_directoryWithFiles, "PrintLayoutPreset.ClearFit.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Report");
                sheet.Cell(1, 1).SetValue("Region");
                sheet.PageSetup.FitToWidth = 1U; sheet.PageSetup.FitToHeight = 0U;
                sheet.PageSetup.ApplyLayout(new ExcelPrintLayoutOptions {
                    Preset = ExcelPrintLayoutPreset.Worksheet
                });
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                Worksheet worksheet = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet;
                PageSetup setup = worksheet.GetFirstChild<PageSetup>()!;
                Assert.Null(setup.FitToWidth);
                Assert.Null(setup.FitToHeight);
                Assert.Equal(100U, setup.Scale!.Value);
                Assert.Null(setup.PaperSize);
                Assert.Null(worksheet.GetFirstChild<SheetProperties>()?.GetFirstChild<PageSetupProperties>()?.FitToPage);
            }
        }

        [Fact]
        public void Test_PageSetupPropertiesWriteAndClearStoredValues() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddSheet("Report");
            ExcelSheetPageSetup page = sheet.PageSetup;

            page.Orientation = OfficePageOrientation.Landscape;
            page.FitToWidth = 1;
            page.FitToHeight = 0;
            Assert.Equal(OfficePageOrientation.Landscape, page.Orientation);
            Assert.Equal(1U, page.FitToWidth);
            Assert.Equal(0U, page.FitToHeight);

            page.Orientation = null;
            page.FitToWidth = null;
            page.FitToHeight = null;
            page.Scale = 80;
            Assert.Null(page.Orientation);
            Assert.Null(page.FitToWidth);
            Assert.Equal(80U, page.Scale);
            Worksheet worksheet = sheet.WorksheetPart.Worksheet;
            Assert.Null(worksheet.GetFirstChild<SheetProperties>()?.GetFirstChild<PageSetupProperties>()?.FitToPage);
            Assert.Null(worksheet.GetFirstChild<PageSetup>()!.Orientation);
        }
}
}

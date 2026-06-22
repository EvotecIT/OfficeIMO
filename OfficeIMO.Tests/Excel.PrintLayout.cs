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
                ExcelSheet sheet = document.AddWorkSheet("Report");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(2, 1, "EU");

                sheet.ApplyPrintLayout(new ExcelPrintLayoutOptions {
                    Preset = ExcelPrintLayoutPreset.Report,
                    PrintArea = "A1:D25",
                    RepeatFirstColumn = 1,
                    RepeatLastColumn = 1
                });
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelSheet sheet = document["Report"];
                ExcelSheetPageSetup setup = sheet.GetPageSetup();
                Assert.Equal(ExcelPageOrientation.Landscape, setup.Orientation);
                Assert.Equal(1U, setup.FitToWidth);
                Assert.Equal(0U, setup.FitToHeight);
                Assert.Equal(ExcelPageOrder.DownThenOver, setup.PageOrder);
                Assert.NotNull(setup.Margins);
                Assert.Equal(0.25D, setup.Margins!.Left);
                Assert.Equal("$A$1:$D$25", sheet.GetPrintArea());

                ExcelPrintTitles titles = sheet.GetPrintTitles();
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
    }
}

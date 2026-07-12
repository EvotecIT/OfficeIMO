using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_WritesWorksheetSelectionMetadata() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorkSheet("Selection");
                    sheet.CellValue(3, 3, "Active");
                    sheet.CellValue(4, 4, "Selected");
                    sheet.CellValue(5, 5, "Selected");

                    Worksheet worksheet = sheet.WorksheetPart.Worksheet;
                    SheetViews? sheetViews = worksheet.GetFirstChild<SheetViews>();
                    if (sheetViews == null) {
                        sheetViews = new SheetViews();
                        worksheet.InsertAt(sheetViews, 0);
                    }

                    SheetView? sheetView = sheetViews.GetFirstChild<SheetView>();
                    if (sheetView == null) {
                        sheetView = new SheetView { WorkbookViewId = 0U };
                        sheetViews.Append(sheetView);
                    }

                    sheetView.RemoveAllChildren<Selection>();
                    sheetView.Append(new Selection {
                        ActiveCell = "C3",
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = "C3 D4:E5" }
                    });
                    worksheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsSelection selection = Assert.Single(legacySheet.Selections);
                Assert.Equal((byte)3, selection.Pane);
                Assert.Equal(3, selection.ActiveRow);
                Assert.Equal(3, selection.ActiveColumn);
                Assert.Equal(new[] { "C3", "D4:E5" }, selection.SelectedRanges.Select(range => range.Reference).ToArray());

                SheetView projectedView = result.Document.Sheets[0]
                    .WorksheetPart
                    .Worksheet
                    .GetFirstChild<SheetViews>()!
                    .GetFirstChild<SheetView>()!;
                Selection projectedSelection = Assert.Single(projectedView.Elements<Selection>());
                Assert.Equal("C3", projectedSelection.ActiveCell!.Value);
                Assert.Equal("C3 D4:E5", projectedSelection.SequenceOfReferences!.InnerText);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }
    }
}

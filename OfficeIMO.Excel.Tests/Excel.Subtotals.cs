using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_SubtotalSummary_WritesFormulasAndOutlinesGroups() {
            string filePath = Path.Combine(_directoryWithFiles, "SubtotalSummary.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Region");
                sheet.Cell(1, 2).SetValue("Sales");
                sheet.Cell(1, 3).SetValue("Units");
                sheet.Cell(2, 1).SetValue("NA");
                sheet.Cell(2, 2).SetValue(100);
                sheet.Cell(2, 3).SetValue(2);
                sheet.Cell(3, 1).SetValue("NA");
                sheet.Cell(3, 2).SetValue(150);
                sheet.Cell(3, 3).SetValue(3);
                sheet.Cell(4, 1).SetValue("EMEA");
                sheet.Cell(4, 2).SetValue(200);
                sheet.Cell(4, 3).SetValue(4);
                sheet.Cell(5, 1).SetValue("EMEA");
                sheet.Cell(5, 2).SetValue(50);
                sheet.Cell(5, 3).SetValue(1);

                ExcelSubtotalResult result = sheet.AddSubtotalSummary(new ExcelSubtotalOptions {
                    DataEndRow = 5,
                    GroupColumn = 1,
                    ValueColumns = new[] { 2, 3 },
                    SummaryStartRow = 7
                });

                Assert.Equal("A7:C10", result.SummaryRange);
                Assert.Equal(2, result.GroupCount);
                Assert.True(result.GrandTotalWritten);
                Assert.Equal("NA", result.Groups[0].Key);
                Assert.Equal(8, result.Groups[0].SummaryRow);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                SheetData sheetData = wsPart.Worksheet.GetFirstChild<SheetData>()!;
                Cell b8 = sheetData.Descendants<Cell>().First(cell => cell.CellReference?.Value == "B8");
                Cell c10 = sheetData.Descendants<Cell>().First(cell => cell.CellReference?.Value == "C10");
                Assert.Equal("SUBTOTAL(9,B2:B3)", b8.CellFormula?.Text);
                Assert.Equal("SUBTOTAL(9,C2:C5)", c10.CellFormula?.Text);

                Row row2 = sheetData.Elements<Row>().First(row => row.RowIndex?.Value == 2U);
                Row row4 = sheetData.Elements<Row>().First(row => row.RowIndex?.Value == 4U);
                Assert.Equal((byte)1, row2.OutlineLevel?.Value);
                Assert.Equal((byte)1, row4.OutlineLevel?.Value);
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_SubtotalSummary_RejectsOutputBeyondWorksheetBounds() {
            string filePath = Path.Combine(_directoryWithFiles, "SubtotalSummaryBounds.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Region");
                sheet.Cell(1, 2).SetValue("Sales");
                sheet.Cell(2, 1).SetValue("NA");
                sheet.Cell(2, 2).SetValue(100);

                Assert.Throws<ArgumentOutOfRangeException>(() => sheet.AddSubtotalSummary(new ExcelSubtotalOptions {
                    DataEndRow = 2,
                    GroupColumn = 1,
                    ValueColumns = new[] { 2 },
                    SummaryStartRow = ExcelA1.MaxRows
                }));

                Assert.Throws<ArgumentOutOfRangeException>(() => sheet.AddSubtotalSummary(new ExcelSubtotalOptions {
                    DataEndRow = 2,
                    GroupColumn = ExcelA1.MaxColumns + 1,
                    ValueColumns = new[] { 2 },
                    SummaryStartRow = 5
                }));
            }
        }
    }
}

using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelFluentWorkbookTests {
        private static string GetCellValue(SpreadsheetDocument document, WorksheetPart worksheetPart, string cellReference) {
            var cell = worksheetPart.Worksheet.Descendants<Cell>().First(c => c.CellReference != null && c.CellReference.Value == cellReference);
            var value = cell.CellValue?.Text ?? string.Empty;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString) {
                var table = document.WorkbookPart.SharedStringTablePart.SharedStringTable;
                if (int.TryParse(value, out int id)) {
                    return table.ChildElements[id].InnerText;
                }
            }
            return value;
        }

        [Fact]
        public void CanBuildWorkbookFluently() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AsFluent()
                    .Sheet("Data", s => s
                        .HeaderRow("Name", "Score")
                        .Row(r => r.Values("Alice", 93))
                        .Row(r => r.Values("Bob", 88))
                        .Table(t => t.Add("A1:B3", true, "Scores"))
                        .Column(c => c.AutoFit()));
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                Assert.Single(document.Sheets);
                var sheetPart = document._spreadSheetDocument.WorkbookPart.WorksheetParts.First();
                Assert.Equal("Name", GetCellValue(document._spreadSheetDocument, sheetPart, "A1"));
                Assert.Equal("93", GetCellValue(document._spreadSheetDocument, sheetPart, "B2"));
                Assert.True(sheetPart.TableDefinitionParts.Any());
            }

            File.Delete(filePath);
        }
    }
}

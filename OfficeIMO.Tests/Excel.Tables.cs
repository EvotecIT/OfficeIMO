using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_AddTableWithStyle() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.SetCellValue(1, 1, "Name");
                sheet.SetCellValue(1, 2, "Value");
                sheet.SetCellValue(2, 1, "A");
                sheet.SetCellValue(2, 2, 1d);
                sheet.AddTable("A1:B2", true, "MyTable", TableStyle.TableStyleMedium9);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                TableDefinitionPart tablePart = wsPart.TableDefinitionParts.First();
                Assert.Equal("A1:B2", tablePart.Table.Reference.Value);
                Assert.Equal("MyTable", tablePart.Table.Name);
                Assert.Equal("TableStyleMedium9", tablePart.Table.TableStyleInfo.Name.Value);
            }
        }
    }
}

using System.Data;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ExcelSheetProtection_TableEditingPreset_AllowsTableWorkflows() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelProtection.TableEditing.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Sales");
                sheet.CellValue(2, 1, "NA");
                sheet.CellValue(2, 2, 100);
                sheet.AddTable("A1:B2", true, "Sales", OfficeIMO.Excel.TableStyle.TableStyleMedium2);
                sheet.ProtectTableEditing();
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = GetWorksheetPartByName(spreadsheet, "Data");
                SheetProtection protection = worksheetPart.Worksheet.Elements<SheetProtection>().Single();
                Assert.True(protection.Sheet!.Value);
                Assert.False(protection.InsertRows!.Value);
                Assert.False(protection.Sort!.Value);
                Assert.False(protection.AutoFilter!.Value);
                Assert.False(protection.SelectUnlockedCells!.Value);
            }
        }

        [Fact]
        public void Test_ExcelSheetProtection_TableEditingPreset_PersistsThroughDirectTableSave() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelProtection.TableEditingDirectSave.xlsx");
            var table = new DataTable("Sales");
            table.Columns.Add("Region", typeof(string));
            table.Columns.Add("Sales", typeof(int));
            table.Rows.Add("NA", 100);
            table.Rows.Add("EMEA", 200);

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.InsertDataTableAsTable(table, includeHeaders: true, tableName: "Sales", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2);
                sheet.ProtectTableEditing();
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = GetWorksheetPartByName(spreadsheet, "Data");
                SheetProtection protection = worksheetPart.Worksheet.Elements<SheetProtection>().Single();
                Assert.True(protection.Sheet!.Value);
                Assert.False(protection.InsertRows!.Value);
                Assert.False(protection.Sort!.Value);
                Assert.False(protection.AutoFilter!.Value);
            }
        }
    }
}

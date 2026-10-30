#region QuickStart
using OfficeIMO.Excel;

using var document = ExcelDocument.Create("report.xlsx");
var sheet = document.AddWorksheet("Data");

sheet.CellValue(1, 1, "Name");
sheet.CellValue(1, 2, "Value");
sheet.CellValue(2, 1, "Alpha");
sheet.CellValue(2, 2, 42);
sheet.AddTable("A1:B2", hasHeader: true, name: "DataTable", style: ExcelTableStyle.TableStyleMedium9);
sheet.AutoFitColumns();

document.Save();
#endregion

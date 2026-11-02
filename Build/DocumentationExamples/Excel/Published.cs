using OfficeIMO.Excel;

using var document = ExcelDocument.Create("report.xlsx");
var sheet = document.AddWorksheet("Data");

sheet.Cell(1, 1).SetValue("Name");
sheet.Cell(1, 2).SetValue("Value");
sheet.Cell(2, 1).SetValue("Alpha");
sheet.Cell(2, 2).SetValue(42);
sheet.AddTable("A1:B2", hasHeader: true, name: "DataTable", style: ExcelTableStyle.TableStyleMedium9);
sheet.AutoFitColumns();

document.Save();

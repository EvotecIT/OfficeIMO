using System;
using System.IO;
using OfficeIMO.Excel;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Excel
{
    internal static partial class WorkbookAuthoring
    {
        public static void Example_Workbook(string folderPath, bool openExcel)
        {
            Console.WriteLine("[*] Excel - Creating workbook with worksheet API");
            string filePath = Path.Combine(folderPath, "WorkbookAuthoring.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath))
            {
                // Set Sequential mode to prevent locking issues with combined operations
                document.Execution.Mode = ExcelExecutionMode.Sequential;
                {
                    var sheet1 = document.AddSheet("Data");
                    int nextRow2 = 1;
                    sheet1.Cell(nextRow2, 1, "Name");
                    sheet1.Cell(nextRow2, 2, "Score");
                    sheet1.Cell(nextRow2, 3, "Department");
                    sheet1.Cell(nextRow2, 4, "Date");
                    sheet1.Cell(nextRow2, 5, "Status");
                    nextRow2++;
                    sheet1.Cell(nextRow2, 1, "Alice");
                    sheet1.Cell(nextRow2, 2, 93);
                    sheet1.Cell(nextRow2, 3, "Engineering");
                    sheet1.Cell(nextRow2, 4, DateTime.Now.AddDays(-10));
                    sheet1.Cell(nextRow2, 5, "Active");
                    nextRow2++;
                    sheet1.Cell(nextRow2, 1, "Bob");
                    sheet1.Cell(nextRow2, 2, 88);
                    sheet1.Cell(nextRow2, 3, "Marketing");
                    sheet1.Cell(nextRow2, 4, DateTime.Now.AddDays(-8));
                    sheet1.Cell(nextRow2, 5, "Active");
                    nextRow2++;
                    sheet1.Cell(nextRow2, 1, "Charlie");
                    sheet1.Cell(nextRow2, 2, 76);
                    sheet1.Cell(nextRow2, 3, "Sales");
                    sheet1.Cell(nextRow2, 4, DateTime.Now.AddDays(-5));
                    sheet1.Cell(nextRow2, 5, "Inactive");
                    nextRow2++;
                    sheet1.Cell(nextRow2, 1, "David");
                    sheet1.Cell(nextRow2, 2, 91);
                    sheet1.Cell(nextRow2, 3, "Engineering");
                    sheet1.Cell(nextRow2, 4, DateTime.Now.AddDays(-3));
                    sheet1.Cell(nextRow2, 5, "Active");
                    nextRow2++;
                    sheet1.Cell(nextRow2, 1, "Eve");
                    sheet1.Cell(nextRow2, 2, 85);
                    sheet1.Cell(nextRow2, 3, "HR");
                    sheet1.Cell(nextRow2, 4, DateTime.Now.AddDays(-2));
                    sheet1.Cell(nextRow2, 5, "Active");
                    nextRow2++;
                    sheet1.Cell(nextRow2, 1, "Frank");
                    sheet1.Cell(nextRow2, 2, 79);
                    sheet1.Cell(nextRow2, 3, "Marketing");
                    sheet1.Cell(nextRow2, 4, DateTime.Now.AddDays(-1));
                    sheet1.Cell(nextRow2, 5, "Active");
                    nextRow2++;
                    sheet1.Cell(nextRow2, 1, "Grace");
                    sheet1.Cell(nextRow2, 2, 94);
                    sheet1.Cell(nextRow2, 3, "Engineering");
                    sheet1.Cell(nextRow2, 4, DateTime.Now);
                    sheet1.Cell(nextRow2, 5, "Active");
                    nextRow2++;
                    sheet1.Cell(nextRow2, 1, "Henry");
                    sheet1.Cell(nextRow2, 2, 82);
                    sheet1.Cell(nextRow2, 3, "Sales");
                    sheet1.Cell(nextRow2, 4, DateTime.Now.AddDays(1));
                    sheet1.Cell(nextRow2, 5, "Inactive");
                    nextRow2++;
                    sheet1.Cell(nextRow2, 1, "Irene");
                    sheet1.Cell(nextRow2, 2, 90);
                    sheet1.Cell(nextRow2, 3, "HR");
                    sheet1.Cell(nextRow2, 4, DateTime.Now.AddDays(2));
                    sheet1.Cell(nextRow2, 5, "Active");
                    nextRow2++;
                    sheet1.Cell(nextRow2, 1, "Jack");
                    sheet1.Cell(nextRow2, 2, 87);
                    sheet1.Cell(nextRow2, 3, "Engineering");
                    sheet1.Cell(nextRow2, 4, DateTime.Now.AddDays(3));
                    sheet1.Cell(nextRow2, 5, "Active");
                    nextRow2++;
                    sheet1.AddTable("A1:E11", true, "EmployeeData");
                    sheet1.Freeze(topRows: 1, leftCols: 1);
                    sheet1.AddConditionalColorScale("B2:B11", Color.Red, Color.Green);
                    sheet1.AddConditionalDataBar("B2:B11", Color.Blue);
                    sheet1.AutoFitColumns();
                    sheet1.AutoFitRows();
                    sheet1.SetColumnWidth(1, ExcelColumnWidth.Characters(15));
                    sheet1.SetColumnWidth(2, ExcelColumnWidth.Characters(10));
                    sheet1.SetColumnWidth(3, ExcelColumnWidth.Characters(15));
                    sheet1.SetColumnWidth(4, ExcelColumnWidth.Characters(12));
                    sheet1.SetColumnWidth(5, ExcelColumnWidth.Characters(10));
                    document.Save();
                } // Don't open Excel yet

                // Validate the document
                var errors = document.DocumentValidationErrors;
                if (errors.Count > 0)
                {
                    Console.WriteLine($"[!] Document has {errors.Count} validation errors:");
                    foreach (var error in errors)
                    {
                        Console.WriteLine($"  - {error.Description}");
                        Console.WriteLine($"    Path: {error.Path ?? "N/A"}");
                        Console.WriteLine($"    Part: {error.PartUri ?? "N/A"}");
                    }
                }
                else
                {
                    Console.WriteLine("[✓] Document is valid");
                }

                if (openExcel)
                {
                    document.OpenInApplication();
                }
            }
        }
    }
}
using System;
using System.IO;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Excel {
    internal static partial class FluentWorkbook {
        public static void Example_FluentWorkbook(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Creating workbook with fluent API");
            string filePath = Path.Combine(folderPath, "FluentWorkbook.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                // Set Sequential mode to prevent locking issues with combined operations
                document.Execution.Mode = ExecutionMode.Sequential;

                document.AsFluent()
                    .Sheet("Data", s => s
                        .HeaderRow("Name", "Score", "Department", "Date", "Status")
                        .Row(r => r.Values("Alice", 93, "Engineering", DateTime.Now.AddDays(-10), "Active"))
                        .Row(r => r.Values("Bob", 88, "Marketing", DateTime.Now.AddDays(-8), "Active"))
                        .Row(r => r.Values("Charlie", 76, "Sales", DateTime.Now.AddDays(-5), "Inactive"))
                        .Row(r => r.Values("David", 91, "Engineering", DateTime.Now.AddDays(-3), "Active"))
                        .Row(r => r.Values("Eve", 85, "HR", DateTime.Now.AddDays(-2), "Active"))
                        .Row(r => r.Values("Frank", 79, "Marketing", DateTime.Now.AddDays(-1), "Active"))
                        .Row(r => r.Values("Grace", 94, "Engineering", DateTime.Now, "Active"))
                        .Row(r => r.Values("Henry", 82, "Sales", DateTime.Now.AddDays(1), "Inactive"))
                        .Row(r => r.Values("Irene", 90, "HR", DateTime.Now.AddDays(2), "Active"))
                        .Row(r => r.Values("Jack", 87, "Engineering", DateTime.Now.AddDays(3), "Active"))
                        .Table(t => t.Add("A1:E11", true, "EmployeeData"))
                        .Freeze(topRows: 1, leftCols: 1)
                        // Note: Table includes AutoFilter by default. Use .WithAutoFilter(false) to disable it.
                        // If you want worksheet-level AutoFilter instead of table AutoFilter, call .AutoFilter() and use .WithAutoFilter(false) on the table
                        //.AutoFilter("A1:E11")
                        .ConditionalColorScale("B2:B11", Color.Red, Color.Green)
                        .ConditionalDataBar("B2:B11", Color.Blue)
                        .AutoFit(columns: true, rows: true)
                        .Columns(c => c
                            .Col(1, col => col.Width(15))
                            .Col(2, col => col.Width(10))
                            .Col(3, col => col.Width(15))
                            .Col(4, col => col.Width(12))
                            .Col(5, col => col.Width(10))
                        )
                    )
                    .End()
                    .Save(false); // Don't open Excel yet

                // Validate the document
                var errors = document.DocumentValidationErrors;
                if (errors.Count > 0) {
                    Console.WriteLine($"[!] Document has {errors.Count} validation errors:");
                    foreach (var error in errors) {
                        Console.WriteLine($"  - {error.Description}");
                        Console.WriteLine($"    Path: {error.Path?.XPath ?? "N/A"}");
                        Console.WriteLine($"    Part: {error.Part?.Uri?.ToString() ?? "N/A"}");
                    }
                } else {
                    Console.WriteLine("[✓] Document is valid");
                }

                if (openExcel) {
                    document.Open();
                }
            }
        }
    }
}

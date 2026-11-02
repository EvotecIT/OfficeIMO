using System;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates detection of overlapping table ranges.
    /// </summary>
    public static class AddTableOverlapping {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Add table overlapping range");
            string filePath = System.IO.Path.Combine(folderPath, "Add Table Overlapping.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Name");
                sheet.Cell(1, 2).SetValue("Value");
                sheet.Cell(2, 1).SetValue("A");
                sheet.Cell(2, 2).SetValue(10d);
                sheet.Cell(3, 1).SetValue("B");
                sheet.Cell(3, 2).SetValue(20d);
                sheet.AddTable("A1:B3", true, "Table1", ExcelTableStyle.TableStyleMedium9);

                try {
                    sheet.AddTable("B2:C4", true, "Table2", ExcelTableStyle.TableStyleMedium9);
                } catch (InvalidOperationException ex) {
                    Console.WriteLine($"Expected error: {ex.Message}");
                }

                document.Save();
                if (openExcel) document.OpenInApplication();
            }
        }
    }
}


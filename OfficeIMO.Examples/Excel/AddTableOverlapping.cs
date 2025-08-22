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
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Value");
                sheet.CellValue(2, 1, "A");
                sheet.CellValue(2, 2, 10d);
                sheet.CellValue(3, 1, "B");
                sheet.CellValue(3, 2, 20d);
                sheet.AddTable("A1:B3", true, "Table1", TableStyle.TableStyleMedium9);

                try {
                    sheet.AddTable("B2:C4", true, "Table2", TableStyle.TableStyleMedium9);
                } catch (InvalidOperationException ex) {
                    Console.WriteLine($"Expected error: {ex.Message}");
                }

                document.Save(openExcel);
            }
        }
    }
}


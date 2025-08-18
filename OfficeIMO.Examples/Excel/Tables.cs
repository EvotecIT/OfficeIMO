using System;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates adding a table with a built-in style.
    /// </summary>
    public static class Tables {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Tables");
            string filePath = System.IO.Path.Combine(folderPath, "Tables.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.SetCellValue(1, 1, "Name");
                sheet.SetCellValue(1, 2, "Value");
                sheet.SetCellValue(2, 1, "A");
                sheet.SetCellValue(2, 2, 10d);
                sheet.SetCellValue(3, 1, "B");
                sheet.SetCellValue(3, 2, 20d);

                sheet.AddTable("A1:B3", true, "MyTable", TableStyle.TableStyleMedium2);

                document.Save(openExcel);
            }
        }
    }
}

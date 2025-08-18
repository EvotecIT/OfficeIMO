using System;
using System.Collections.Generic;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates adding an autofilter to a worksheet.
    /// </summary>
    public static class AutoFilter {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - AutoFilter example");
            string filePath = System.IO.Path.Combine(folderPath, "AutoFilter.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.SetCellValue(1, 1, "Name");
                sheet.SetCellValue(1, 2, "Value");
                sheet.SetCellValue(2, 1, "A");
                sheet.SetCellValue(2, 2, 10d);
                sheet.SetCellValue(3, 1, "B");
                sheet.SetCellValue(3, 2, 20d);

                Dictionary<uint, IEnumerable<string>> criteria = new Dictionary<uint, IEnumerable<string>> {
                    { 0, new[] { "A" } }
                };

                sheet.AddAutoFilter("A1:B3", criteria);
                document.Save(openExcel);
            }
        }
    }
}

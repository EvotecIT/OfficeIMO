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
                ExcelSheet sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Name");
                sheet.Cell(1, 2).SetValue("Value");
                sheet.Cell(2, 1).SetValue("A");
                sheet.Cell(2, 2).SetValue(10d);
                sheet.Cell(3, 1).SetValue("B");
                sheet.Cell(3, 2).SetValue(20d);

                Dictionary<uint, IEnumerable<string>> criteria = new Dictionary<uint, IEnumerable<string>> {
                    { 0, new[] { "A" } }
                };

                sheet.Range("A1:B3").AutoFilter();
                foreach (var criterion in criteria) sheet.Range("A1:B3").AutoFilter().Column(criterion.Key).Equals(criterion.Value);
                document.Save();
                if (openExcel) document.OpenInApplication();
            }
        }
    }
}

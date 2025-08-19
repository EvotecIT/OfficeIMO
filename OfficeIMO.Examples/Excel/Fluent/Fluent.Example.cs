using System;
using System.IO;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;

namespace OfficeIMO.Examples.Excel {
    internal static partial class FluentWorkbook {
        public static void Example_FluentWorkbook(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Creating workbook with fluent API");
            string filePath = Path.Combine(folderPath, "FluentWorkbook.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AsFluent()
                    .Sheet("Data", s => s
                        .HeaderRow("Name", "Score")
                        .Row(r => r.Values("Alice", 93))
                        .Row(r => r.Values("Bob", 88))
                        .Table(t => t.Add("A1:B3", true, "Scores"))
                        .Column(c => c.AutoFit()));
                document.Save(openExcel);
            }
        }
    }
}

using System;
using System.IO;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;

namespace OfficeIMO.Examples.Excel {
    internal static partial class FluentWorkbook {
        public static void Example_FluentCells(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Fluent cell helpers");
            string filePath = Path.Combine(folderPath, "FluentCells.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AsFluent()
                    .Sheet("Data", s => s
                        .Cell("B2", "Direct")
                        .Row(r => r.Cell("C", "Row builder"))
                        .Range("A1:C3", r => r.Cell(3, 3, "Range cell")))
                    .End()
                    .Save(openExcel);
            }
        }
    }
}

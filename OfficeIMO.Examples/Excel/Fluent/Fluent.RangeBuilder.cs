using System;
using System.IO;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;

namespace OfficeIMO.Examples.Excel {
    internal static partial class FluentWorkbook {
        public static void Example_RangeBuilder(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Using RangeBuilder");
            string filePath = Path.Combine(folderPath, "FluentRangeBuilder.xlsx");
            object[,] values = { { "A", "B" }, { "C", "D" } };
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AsFluent()
                    .Sheet("Data", s => s.Range("A1:B2", r => {
                        r.Set(values);
                        r.NumberFormat("@");
                    }))
                    .End()
                    .Save(openExcel);
            }
        }
    }
}


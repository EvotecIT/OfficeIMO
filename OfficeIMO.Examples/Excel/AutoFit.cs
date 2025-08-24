using System;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates autofitting columns and rows when writing cell values.
    /// </summary>
    public static class AutoFit {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - AutoFit columns and rows");
            string filePath = System.IO.Path.Combine(folderPath, "AutoFit.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                // Sheet 1: Mixed simple texts
                var s1 = document.AddWorkSheet("Simple");
                s1.CellValue(1, 1, "This is a very long piece of text");
                s1.CellValue(1, 2, "Short");
                s1.CellValue(2, 1, "Second line\nwith newline");
                s1.CellValue(3, 1, "Line1\nLine2\nLine3");
                s1.AutoFitColumns();
                s1.AutoFitRows();

                // Sheet 2: Large fonts and italic/bold to check style impact
                var s2 = document.AddWorkSheet("Styled");
                s2.Cell(1, 1, value: "Bold text sample", numberFormat: null);
                s2.Cell(2, 1, value: "Italic sample");
                s2.Cell(3, 1, value: "Bold Italic with\nnewline");
                // Apply some styles via number format call to ensure style part exists
                s2.FormatCell(1, 1, "@");
                s2.FormatCell(2, 1, "@");
                s2.FormatCell(3, 1, "@");
                s2.AutoFitColumns();
                s2.AutoFitRows();

                // Sheet 3: Very long words (no spaces)
                var s3 = document.AddWorkSheet("LongWords");
                s3.CellValue(1, 1, new string('A', 60));
                s3.CellValue(2, 1, new string('B', 40));
                s3.AutoFitColumns();
                s3.AutoFitRows();

                document.Save(openExcel);
            }
        }
    }
}

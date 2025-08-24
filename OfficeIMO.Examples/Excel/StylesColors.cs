using System;
using System.IO;
using OfficeIMO.Excel;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Excel
{
    internal static class StylesColors
    {
        public static void Example(string folderPath, bool openExcel)
        {
            Console.WriteLine("[*] Excel - Styles & Colors (cells, headers)");
            string filePath = Path.Combine(folderPath, "Excel-StylesColors.xlsx");

            using (var doc = ExcelDocument.Create(filePath, "Data"))
            {
                var s = doc[0];
                // Headers
                s.CellValue(1, 1, "Name");
                s.CellValue(1, 2, "Value");
                s.CellValue(1, 3, "Status");
                // Data
                s.CellValue(2, 1, "Alpha");
                s.CellValue(2, 2, 10);
                s.CellValue(2, 3, "New");

                s.CellValue(3, 1, "Beta");
                s.CellValue(3, 2, 120.5);
                s.CellValue(3, 3, "Hold");

                // Header styling via builder (background + bold)
                s.ColumnStyleByHeader("Name", includeHeader: true)
                 .Background(Color.Parse("#E6EEF8"))
                 .Bold();
                s.ColumnStyleByHeader("Value", includeHeader: true)
                 .Number(decimals: 2)
                 .Background("#FFF3CD")
                 .Bold();
                s.ColumnStyleByHeader("Status", includeHeader: true)
                 .Background(Color.Parse("#E7FFE7"))
                 .Bold();

                // Cell-level background using SixLabors color
                s.CellBackground(2, 3, Color.Parse("#FFFBE6"));
                // Cell-level background using hex string
                s.CellBackground(3, 3, "#FFE7E7");

                // Auto-fit for nice layout
                s.AutoFitColumns();
                s.AutoFitRows();

                doc.Save(openExcel);
            }
        }
    }
}


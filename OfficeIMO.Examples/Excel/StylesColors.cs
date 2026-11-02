using System;
using System.IO;
using OfficeIMO.Excel;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Excel
{
    internal static class StylesColors
    {
        public static void Example(string folderPath, bool openExcel)
        {
            Console.WriteLine("[*] Excel - Styles & Colors (cells, headers)");
            string filePath = Path.Combine(folderPath, "Excel-StylesColors.xlsx");

            using (var doc = ExcelDocument.Create(filePath, new ExcelCreateOptions { FirstSheetName = "Data" }))
            {
                var s = doc[0];
                // Headers
                s.Cell(1, 1).SetValue("Name");
                s.Cell(1, 2).SetValue("Value");
                s.Cell(1, 3).SetValue("Status");
                // Data
                s.Cell(2, 1).SetValue("Alpha");
                s.Cell(2, 2).SetValue(10);
                s.Cell(2, 3).SetValue("New");

                s.Cell(3, 1).SetValue("Beta");
                s.Cell(3, 2).SetValue(120.5);
                s.Cell(3, 3).SetValue("Hold");

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

                // Cell-level background using OfficeIMO color
                s.Cell(2, 3).SetFillColor(Color.Parse("#FFFBE6"));
                // Cell-level background using hex string
                s.Cell(3, 3).SetFillColor("#FFE7E7");

                // Auto-fit for nice layout
                s.AutoFitColumns();
                s.AutoFitRows();

                doc.Save();
                if (openExcel) doc.OpenInApplication();
            }
        }
    }
}


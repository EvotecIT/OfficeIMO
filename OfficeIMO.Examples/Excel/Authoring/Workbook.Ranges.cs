using System;
using System.IO;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel
{
    internal static partial class WorkbookAuthoring
    {
        public static void Example_Ranges(string folderPath, bool openExcel)
        {
            Console.WriteLine("[*] Excel - Using Ranges");
            string filePath = Path.Combine(folderPath, "FluentRanges.xlsx");
            object[, ] values =
            {
                {
                    "A",
                    "B"
                },
                {
                    "C",
                    "D"
                }
            };
            using (ExcelDocument document = ExcelDocument.Create(filePath))
            {
                {
                    var sheet1 = document.AddWorksheet("Data");

                    var range3 = sheet1.Range("A1:B2");
                    range3.SetValues(values);
                    range3.SetNumberFormat("@");
                    document.Save();
                }

                if (openExcel)
                    document.OpenInApplication();
            }
        }
    }
}
using System;
using System.IO;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel
{
    internal static partial class WorkbookAuthoring
    {
        public static void Example_Cells(string folderPath, bool openExcel)
        {
            Console.WriteLine("[*] Excel - Cell helpers");
            string filePath = Path.Combine(folderPath, "FluentCells.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath))
            {
                {
                    var sheet1 = document.AddSheet("Data");
                    int nextRow2 = 1;
                    {
                        var cell1 = sheet1.CellAt("B2");
                        sheet1.Cell(cell1.Row, cell1.Column, "Direct");
                    }

                    {
                        var cell2 = sheet1.CellAt("C" + nextRow2);
                        sheet1.Cell(cell2.Row, cell2.Column, "Row");
                    }

                    nextRow2++;
                    var range3 = sheet1.Range("A1:C3");
                    sheet1.Cell(range3.CellAt(3, 3).Row, range3.CellAt(3, 3).Column, "Range cell");
                    document.Save();
                }

                if (openExcel)
                    document.OpenInApplication();
            }
        }
    }
}
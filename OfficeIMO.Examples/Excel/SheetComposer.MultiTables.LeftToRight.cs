using System;
using System.Collections.Generic;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;

namespace OfficeIMO.Examples.Excel
{
    internal static class SheetComposerMultiTables
    {
        public static void Example_LeftToRight(string folderPath, bool openExcel)
        {
            string filePath = System.IO.Path.Combine(folderPath, "Excel.SheetComposer.LeftToRight.Tables.xlsx");

            using var doc = ExcelDocument.Create(filePath);
            // Ensure predictable behavior for mixed operations
            doc.Execution.Mode = ExecutionMode.Sequential;

            var left = new List<object> {
                new { Name = "Alpha", Value = 1, Note = "short" },
                new { Name = "Beta", Value = 2, Note = "wrap\nthis" },
                new { Name = "Gamma", Value = 3, Note = "ok" }
            };
            var middle = new List<object> {
                new { Key = "K1", Count = 12 }, new { Key = "K2", Count = 3 }, new { Key = "K3", Count = 8 }
            };
            var right = new List<object> {
                new { Title = "Link 1", Url = "https://example.com/1" },
                new { Title = "Link 2", Url = "https://example.com/2" },
            };

            var s = new SheetComposer(doc, "Left→Right");
            s.Title("Tables (Left → Right)", "Each table starts on row 3 and flows horizontally.");

            // At current row (3), place 3 columns of width 3 with a 1-column gutter
            // This yields blocks starting at columns 1, 5, and 9 (1 + 3 + 1).
            s.Columns(3, cols =>
            {
                // Column A
                cols[0].Section("A");
                cols[0].TableFrom(left, title: null, visuals: v => { v.FreezeHeaderRow = false; });

                // Column O (A + 12 + 2)
                cols[1].Section("B");
                cols[1].TableFrom(middle, title: null, visuals: v => { v.FreezeHeaderRow = false; });

                // Column AC
                cols[2].Section("C");
                var range = cols[2].TableFrom(right, title: null, visuals: v => { v.FreezeHeaderRow = false; });
                s.ApplyColumnSizing(range, opt => {
                    opt.MediumHeaders.Add("Title");
                    opt.LongHeaders.Add("Url");
                    opt.WrapHeaders.Add("Title");
                });
            }, columnWidth: 3, gutter: 1);

            s.Finish(autoFitColumns: false);
            doc.Save(openExcel);
        }
    }
}

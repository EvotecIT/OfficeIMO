using System;
using System.Data;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates inserting a DataTable that includes TimeSpan values.
    /// </summary>
    public static class DataTableTimeSpans {
        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - DataTable with TimeSpan durations");
            string filePath = System.IO.Path.Combine(folderPath, "DataTable TimeSpans.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Durations");

                var table = new DataTable();
                table.Columns.Add("Task", typeof(string));
                table.Columns.Add("Duration", typeof(TimeSpan));

                table.Rows.Add("Planning", TimeSpan.FromMinutes(45));
                table.Rows.Add("Implementation", TimeSpan.FromHours(1.5));
                table.Rows.Add("Testing", new TimeSpan(2, 10, 0));

                sheet.InsertDataTable(table, includeHeaders: true);

                document.Save(openExcel);
            }
        }
    }
}

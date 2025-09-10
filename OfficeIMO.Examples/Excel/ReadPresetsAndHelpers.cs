using System;
using System.Data;
using System.IO;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel
{
    /// <summary>
    /// Demonstrates using read presets and static helpers for easy reading.
    /// </summary>
    internal static class ReadPresetsAndHelpers
    {
        private sealed class SimpleSale
        {
            public bool Active { get; set; }
            public decimal Amount { get; set; }
            public DateTime Date { get; set; }
            public int Qty { get; set; }
            public string Note { get; set; }
        }

        public static void Example(string folderPath, bool openExcel)
        {
            Console.WriteLine("[*] Excel - Read presets and helpers");
            string filePath = Path.Combine(folderPath, "Read-Presets.xlsx");

            // 1) Create a tiny workbook to read back
            using (var doc = ExcelDocument.Create(filePath, "Data"))
            {
                var sheet = doc.Sheets[0];
                sheet.CellValue(1, 1, "Active");
                sheet.CellValue(1, 2, "Amount");
                sheet.CellValue(1, 3, "Date");
                sheet.CellValue(1, 4, "Qty");
                sheet.CellValue(1, 5, "Note");

                sheet.CellValue(2, 1, "Yes");
                sheet.CellValue(2, 2, "$999.95");
                sheet.CellValue(2, 3, DateTime.Today);
                sheet.CellValue(2, 4, 2);
                sheet.CellValue(2, 5, "Preset demo");

                sheet.CellValue(3, 1, "No");
                sheet.CellValue(3, 2, "1 234,56 €");
                sheet.CellValue(3, 3, DateTime.Today.AddDays(-2));
                sheet.CellValue(3, 4, 5);
                sheet.CellValue(3, 5, "Helpers demo");

                doc.Save(openExcel: openExcel);
            }

            // 2) Use presets with the standard reader
            var simple = ExcelReadPresets.Simple();
            using (var rdr = ExcelDocumentReader.Open(filePath, simple))
            {
                var sheet = rdr.GetSheet("Data");
                var items = sheet.ReadObjects<SimpleSale>("A1:E3");
                foreach (var x in items)
                {
                    Console.WriteLine($"Simple preset → Active={x.Active}, Amount={x.Amount}, Date={x.Date:d}, Qty={x.Qty}, Note={x.Note}");
                }
            }

            // 3) Use static helpers for one-liners
            var aggressive = ExcelReadPresets.Aggressive();
            DataTable dt = ExcelRead.ReadRangeAsDataTable(filePath, "Data", "A1:E3", headersInFirstRow: true, options: ExcelReadPresets.None());
            Console.WriteLine($"Helpers ReadTable (None preset): rows={dt.Rows.Count}, cols={dt.Columns.Count}");

            var list = ExcelRead.ReadRangeObjectsAs<SimpleSale>(filePath, "Data", "A1:E3", options: aggressive);
            Console.WriteLine($"Helpers ReadObjects (Aggressive preset): count={list.Count}");

            var amounts = ExcelRead.ReadColumnAs<decimal>(filePath, "Data", "B2:B3", options: ExcelReadPresets.Simple());
            foreach (var a in amounts)
                Console.WriteLine($"Helpers ReadColumnAs<decimal>: {a}");
        }
    }
}

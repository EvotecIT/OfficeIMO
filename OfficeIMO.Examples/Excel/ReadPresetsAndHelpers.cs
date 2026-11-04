using System;
using System.IO;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel
{
    /// <summary>
    /// Demonstrates the canonical OfficeIMO.Excel data reader.
    /// </summary>
    internal static class ReadPresetsAndHelpers
    {
        private sealed class SimpleSale
        {
            public bool Active { get; set; }
            public decimal Amount { get; set; }
            public DateTime Date { get; set; }
            public int Qty { get; set; }
            public string Note { get; set; } = string.Empty;
        }

        public static void Example(string folderPath, bool openExcel)
        {
            Console.WriteLine("[*] Excel - Read presets and helpers");
            string filePath = Path.Combine(folderPath, "Read-Presets.xlsx");

            // 1) Create a tiny workbook to read back
            using (var doc = ExcelDocument.Create(filePath, new ExcelCreateOptions { FirstSheetName = "Data" }))
            {
                var sheet = doc.Sheets[0];
                sheet.Cell(1, 1).SetValue("Active");
                sheet.Cell(1, 2).SetValue("Amount");
                sheet.Cell(1, 3).SetValue("Date");
                sheet.Cell(1, 4).SetValue("Qty");
                sheet.Cell(1, 5).SetValue("Note");

                sheet.Cell(2, 1).SetValue(true);
                sheet.Cell(2, 2).SetValue(999.95m);
                sheet.Cell(2, 3).SetValue(DateTime.Today);
                sheet.Cell(2, 4).SetValue(2);
                sheet.Cell(2, 5).SetValue("Preset demo");

                sheet.Cell(3, 1).SetValue(false);
                sheet.Cell(3, 2).SetValue(1234.56m);
                sheet.Cell(3, 3).SetValue(DateTime.Today.AddDays(-2));
                sheet.Cell(3, 4).SetValue(5);
                sheet.Cell(3, 5).SetValue("Helpers demo");

                doc.Save();
                if (openExcel) doc.OpenInApplication();
            }

            // 2) One reader shape, automatic used range, and typed getters.
            using var reader = ExcelDocument.OpenReader(filePath, new ExcelReaderOptions {
                NumericAsDecimal = true
            });
            while (reader.Read()) {
                var item = new SimpleSale {
                    Active = reader.GetBoolean(reader.GetOrdinal("Active")),
                    Amount = reader.GetDecimal(reader.GetOrdinal("Amount")),
                    Date = reader.GetDateTime(reader.GetOrdinal("Date")),
                    Qty = reader.GetInt32(reader.GetOrdinal("Qty")),
                    Note = reader.GetString(reader.GetOrdinal("Note"))
                };
                Console.WriteLine($"Active={item.Active}, Amount={item.Amount}, Date={item.Date:d}, Qty={item.Qty}, Note={item.Note}");
            }
        }
    }
}

using System;
using System.IO;
using OfficeIMO.Excel;
using OfficeIMO.Tabular;

namespace OfficeIMO.Examples.Excel
{
    /// <summary>
    /// Demonstrates the canonical format-neutral tabular reader.
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
            using (var doc = ExcelDocument.Create(filePath, "Data"))
            {
                var sheet = doc.Sheets[0];
                sheet.CellValue(1, 1, "Active");
                sheet.CellValue(1, 2, "Amount");
                sheet.CellValue(1, 3, "Date");
                sheet.CellValue(1, 4, "Qty");
                sheet.CellValue(1, 5, "Note");

                sheet.CellValue(2, 1, true);
                sheet.CellValue(2, 2, 999.95m);
                sheet.CellValue(2, 3, DateTime.Today);
                sheet.CellValue(2, 4, 2);
                sheet.CellValue(2, 5, "Preset demo");

                sheet.CellValue(3, 1, false);
                sheet.CellValue(3, 2, 1234.56m);
                sheet.CellValue(3, 3, DateTime.Today.AddDays(-2));
                sheet.CellValue(3, 4, 5);
                sheet.CellValue(3, 5, "Helpers demo");

                doc.Save();
                if (openExcel) doc.OpenInApplication();
            }

            // 2) One reader shape, automatic used range, and typed getters.
            using var reader = TabularReader.Open(filePath, new TabularReadOptions {
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

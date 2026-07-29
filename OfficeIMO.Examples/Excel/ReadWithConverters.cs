using System;
using System.IO;
using OfficeIMO.Excel;
using OfficeIMO.Tabular;

namespace OfficeIMO.Examples.Excel
{
    internal static class ReadWithConverters
    {
        private sealed class Sale
        {
            public bool Active { get; set; }
            public decimal Amount { get; set; }
            public DateTime Date { get; set; }
            public int Qty { get; set; }
            public string Note { get; set; } = string.Empty;
        }

        public static void Example(string folderPath, bool openExcel)
        {
            string filePath = Path.Combine(folderPath, "Read-With-Converters.xlsx");

            // 1) Create a small workbook to read back
            using (var doc = ExcelDocument.Create(filePath, "Data"))
            {
                var sheet = doc.Sheets[0];
                sheet.CellValue(1, 1, "Active");
                sheet.CellValue(1, 2, "Amount");
                sheet.CellValue(1, 3, "Date");
                sheet.CellValue(1, 4, "Qty");
                sheet.CellValue(1, 5, "Note");

                // Data rows
                sheet.CellValue(2, 1, true);
                sheet.CellValue(2, 2, 1234.56m);
                sheet.CellValue(2, 3, DateTime.Today);
                sheet.CellValue(2, 4, 3);
                sheet.CellValue(2, 5, "First order");

                sheet.CellValue(3, 1, false);
                sheet.CellValue(3, 2, 2345.67m);
                sheet.CellValue(3, 3, DateTime.Today.AddDays(-1));
                sheet.CellValue(3, 4, 7);
                sheet.CellValue(3, 5, "Second order");

                doc.Save();
                if (openExcel) doc.OpenInApplication();
            }

            // 2) Read the workbook through the same API used for CSV and XLSB.
            using var reader = TabularReader.Open(filePath, new TabularReadOptions { NumericAsDecimal = true });
            while (reader.Read()) {
                var sale = new Sale {
                    Active = reader.GetBoolean(reader.GetOrdinal("Active")),
                    Amount = reader.GetDecimal(reader.GetOrdinal("Amount")),
                    Date = reader.GetDateTime(reader.GetOrdinal("Date")),
                    Qty = reader.GetInt32(reader.GetOrdinal("Qty")),
                    Note = reader.GetString(reader.GetOrdinal("Note"))
                };
                Console.WriteLine($"Active={sale.Active}, Amount={sale.Amount}, Date={sale.Date:d}, Qty={sale.Qty}, Note={sale.Note}");
            }
        }
    }
}

using System;
using System.IO;
using OfficeIMO.Excel;

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
            using (var doc = ExcelDocument.Create(filePath, new ExcelCreateOptions { FirstSheetName = "Data" }))
            {
                var sheet = doc.Sheets[0];
                sheet.Cell(1, 1).SetValue("Active");
                sheet.Cell(1, 2).SetValue("Amount");
                sheet.Cell(1, 3).SetValue("Date");
                sheet.Cell(1, 4).SetValue("Qty");
                sheet.Cell(1, 5).SetValue("Note");

                // Data rows
                sheet.Cell(2, 1).SetValue(true);
                sheet.Cell(2, 2).SetValue(1234.56m);
                sheet.Cell(2, 3).SetValue(DateTime.Today);
                sheet.Cell(2, 4).SetValue(3);
                sheet.Cell(2, 5).SetValue("First order");

                sheet.Cell(3, 1).SetValue(false);
                sheet.Cell(3, 2).SetValue(2345.67m);
                sheet.Cell(3, 3).SetValue(DateTime.Today.AddDays(-1));
                sheet.Cell(3, 4).SetValue(7);
                sheet.Cell(3, 5).SetValue("Second order");

                doc.Save();
                if (openExcel) doc.OpenInApplication();
            }

            // 2) Read the workbook through the package-owned XLSX/XLSB data-reader API.
            using var reader = ExcelDocument.OpenReader(
                filePath,
                new ExcelReaderOptions { NumericAsDecimal = true });
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

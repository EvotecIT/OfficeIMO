using System;
using System.Data;
using System.Globalization;
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
            using (var doc = ExcelDocument.Create(filePath, "Data"))
            {
                var sheet = doc.Sheets[0];
                sheet.CellValue(1, 1, "Active");
                sheet.CellValue(1, 2, "Amount");
                sheet.CellValue(1, 3, "Date");
                sheet.CellValue(1, 4, "Qty");
                sheet.CellValue(1, 5, "Note");

                // Data rows
                sheet.CellValue(2, 1, "Y");
                sheet.CellValue(2, 2, "$1,234.56");
                sheet.CellValue(2, 3, DateTime.Today);
                sheet.CellValue(2, 4, 3);
                sheet.CellValue(2, 5, "First order");

                sheet.CellValue(3, 1, "N");
                sheet.CellValue(3, 2, "PLN 2 345,67");
                sheet.CellValue(3, 3, DateTime.Today.AddDays(-1));
                sheet.CellValue(3, 4, 7);
                sheet.CellValue(3, 5, "Second order");

                doc.Save(openExcel: openExcel);
            }

            // 2) Use a simple preset so the example stays short
            var readOptions = ExcelReadPresets.Simple();

            // 3) Read the file using the reader API
            using (var reader = ExcelDocumentReader.Open(filePath, readOptions))
            {
                var sheet = reader.GetSheet("Data");

                // Read as DataTable
                DataTable dt = sheet.ReadRangeAsDataTable("A1:E3", headersInFirstRow: true);
                Console.WriteLine($"DataTable rows: {dt.Rows.Count}, cols: {dt.Columns.Count}");

                // Map to objects
                foreach (var sale in sheet.ReadObjects<Sale>("A1:E3"))
                {
                    Console.WriteLine($"Active={sale.Active}, Amount={sale.Amount}, Date={sale.Date:d}, Qty={sale.Qty}, Note={sale.Note}");
                }

                // Typed column
                foreach (var amount in sheet.ReadColumnAs<decimal>("B2:B3"))
                {
                    Console.WriteLine($"Amount column value: {amount}");
                }
            }
        }
    }
}

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Reader_ExcelNumericArtifacts_UseStoredDoubleSemanticsAcrossPublicReadPaths() {
            const string storedNumber = "165258.23999999999";
            const decimal expectedExcelNumber = 165258.24m;
            const decimal expectedTextNumber = 165258.23999999999m;
            string filePath = Path.Combine(_directoryWithFiles, "Reader.ExcelNumericArtifacts.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Amount");
                sheet.CellValue(1, 2, "TextAmount");
                sheet.CellValue(2, 1, 1d);
                sheet.CellValue(2, 2, storedNumber);
                sheet.CellValue(4098, 1, 1d);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet;
                var amountCell = worksheet.Descendants<Cell>().Single(cell => cell.CellReference == "A2");
                amountCell.DataType = CellValues.Number;
                amountCell.CellValue = new CellValue(storedNumber);
                worksheet.Save();
            }

            using (var reader = ExcelDocumentReader.Open(filePath)) {
                var materialized = Assert.Single(reader.GetSheet("Data").ReadObjects<DecimalArtifactRow>("A1:B2"));
                Assert.Equal(expectedExcelNumber, materialized.Amount);
                Assert.Equal(expectedTextNumber, materialized.TextAmount);
            }

            using (var reader = ExcelDocumentReader.Open(filePath)) {
                var streamed = Assert.Single(reader.GetSheet("Data").ReadObjectsStream<DecimalArtifactRow>("A1:B2"));
                Assert.Equal(expectedExcelNumber, streamed.Amount);
                Assert.Equal(expectedTextNumber, streamed.TextAmount);
            }

            using (var document = ExcelDocument.Load(
                filePath,
                new ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                var row = Assert.Single(document.GetSheet("Data").RowsAs<DecimalArtifactRow>("A1:B2"));
                Assert.Equal(expectedExcelNumber, row.Amount);
                Assert.Equal(expectedTextNumber, row.TextAmount);
            }

            var decimalOptions = new ExcelReadOptions { NumericAsDecimal = true };
            using (var reader = ExcelDocumentReader.Open(filePath, decimalOptions)) {
                object?[,] values = reader.GetSheet("Data").ReadRange("A2:B2", ExecutionMode.Sequential);
                Assert.Equal(expectedExcelNumber, Assert.IsType<decimal>(values[0, 0]));
                Assert.Equal(storedNumber, Assert.IsType<string>(values[0, 1]));
            }

            using (var reader = ExcelDocumentReader.Open(filePath, decimalOptions))
            using (var dataReader = reader.GetSheet("Data").ReadRangeAsDataReader("A1:B4098", schemaSampleRows: 0)) {
                Assert.True(dataReader.Read());
                Assert.Equal(expectedExcelNumber, Assert.IsType<decimal>(dataReader.GetValue(0)));
                Assert.Equal(storedNumber, dataReader.GetString(1));
            }
        }

        private sealed class DecimalArtifactRow {
            public decimal Amount { get; set; }

            public decimal TextAmount { get; set; }
        }
    }
}

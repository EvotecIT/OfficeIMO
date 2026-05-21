using System.Data;
using System.IO;
using System.Linq;
using System.Text.Json;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_DataExchange_CsvRoundTripPreservesQuotedFields() {
            string filePath = Path.Combine(_directoryWithFiles, "DataExchange.Csv.xlsx");
            const string csv = "Name,Note,Amount\r\nAlpha,\"Hello, \"\"world\"\"\",10.5\r\nBeta,\"Line\r\nbreak\",20\r\n";

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                string range = sheet.FromCsv(csv);
                Assert.Equal("A1:C3", range);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelSheet sheet = document.GetSheet("Data");
                DataTable table = sheet.ToDataTable("A1:C3");

                Assert.Equal(new[] { "Name", "Note", "Amount" }, table.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray());
                Assert.Equal("Alpha", table.Rows[0]["Name"]);
                Assert.Equal("Hello, \"world\"", table.Rows[0]["Note"]);
                Assert.Equal("Line\nbreak", table.Rows[1]["Note"]);

                string exported = sheet.ToCsv("A1:C3");
                Assert.Contains("\"Hello, \"\"world\"\"\"", exported);
                Assert.Contains("\"Line\nbreak\"", exported);
            }
        }

        [Fact]
        public void Test_DataExchange_JsonRoundTripUsesHeaderObjects() {
            string filePath = Path.Combine(_directoryWithFiles, "DataExchange.Json.xlsx");
            const string json = "[{\"Name\":\"Alpha\",\"Amount\":10,\"Active\":true},{\"Name\":\"Beta\",\"Amount\":20.5,\"Active\":false}]";

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                string range = sheet.FromJson(json, startRow: 2, startColumn: 2);
                Assert.Equal("B2:D4", range);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelSheet sheet = document.GetSheet("Data");
                string exported = sheet.ToJson("B2:D4");

                using JsonDocument parsed = JsonDocument.Parse(exported);
                JsonElement rows = parsed.RootElement;
                Assert.Equal(JsonValueKind.Array, rows.ValueKind);
                Assert.Equal(2, rows.GetArrayLength());
                Assert.Equal("Alpha", rows[0].GetProperty("Name").GetString());
                Assert.Equal(10, rows[0].GetProperty("Amount").GetInt32());
                Assert.True(rows[0].GetProperty("Active").GetBoolean());
                Assert.Equal("Beta", rows[1].GetProperty("Name").GetString());
            }
        }
    }
}

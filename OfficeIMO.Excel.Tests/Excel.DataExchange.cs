using System.Data;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_DataExchange_JsonRoundTripUsesHeaderObjects() {
            string filePath = Path.Combine(_directoryWithFiles, "DataExchange.Json.xlsx");
            const string json = "[{\"Name\":\"Alpha\",\"Amount\":10,\"Active\":true},{\"Name\":\"Beta\",\"Amount\":20.5,\"Active\":false}]";

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                string range = sheet.FromJson(json, startRow: 2, startColumn: 2);
                Assert.Equal("B2:D4", range);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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

        [Fact]
        public void Test_DataExchange_JsonCustomOptionsStillApply() {
            string filePath = Path.Combine(_directoryWithFiles, "DataExchange.JsonOptions.xlsx");
            const string json = "[{\"FirstName\":\"Alpha\",\"Amount\":42}]";

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.FromJson(json);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelSheet sheet = document.GetSheet("Data");
                string exported = sheet.ToJson("A1:B2", jsonOptions: new JsonSerializerOptions {
                    DictionaryKeyPolicy = JsonNamingPolicy.CamelCase,
                    NumberHandling = JsonNumberHandling.WriteAsString,
                    Converters = { new UppercaseStringConverter() }
                });

                using JsonDocument parsed = JsonDocument.Parse(exported);
                Assert.True(parsed.RootElement[0].TryGetProperty("firstName", out JsonElement firstName));
                Assert.Equal("ALPHA", firstName.GetString());
                Assert.Equal("42", parsed.RootElement[0].GetProperty("amount").GetString());
            }
        }

        [Fact]
        public void Test_DataExchange_JsonRowConverterStillAppliesWithoutWholeTableMaterialization() {
            string filePath = Path.Combine(_directoryWithFiles, "DataExchange.JsonRowConverter.xlsx");
            const string json = "[{\"Name\":\"Alpha\"}]";

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.FromJson(json);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelSheet sheet = document.GetSheet("Data");
                var options = new JsonSerializerOptions();
                options.Converters.Add(new MarkedRowConverter());

                using JsonDocument parsed = JsonDocument.Parse(sheet.ToJson("A1:A2", jsonOptions: options));

                Assert.True(parsed.RootElement[0].GetProperty("converted").GetBoolean());
                Assert.Equal("Alpha", parsed.RootElement[0].GetProperty("Name").GetString());
            }
        }

        private sealed class UppercaseStringConverter : JsonConverter<string> {
            public override string? Read(ref Utf8JsonReader reader, System.Type typeToConvert, JsonSerializerOptions options) =>
                reader.GetString();

            public override void Write(Utf8JsonWriter writer, string value, JsonSerializerOptions options) =>
                writer.WriteStringValue(value.ToUpperInvariant());
        }

        private sealed class MarkedRowConverter : JsonConverter<Dictionary<string, object?>> {
            public override Dictionary<string, object?>? Read(ref Utf8JsonReader reader, System.Type typeToConvert, JsonSerializerOptions options) =>
                throw new System.NotSupportedException();

            public override void Write(Utf8JsonWriter writer, Dictionary<string, object?> value, JsonSerializerOptions options) {
                writer.WriteStartObject();
                writer.WriteBoolean("converted", true);
                foreach (KeyValuePair<string, object?> property in value) {
                    writer.WritePropertyName(property.Key);
                    if (property.Value == null) {
                        writer.WriteNullValue();
                    } else {
                        JsonSerializer.Serialize(writer, property.Value, property.Value.GetType(), options);
                    }
                }
                writer.WriteEndObject();
            }
        }

        [Fact]
        public void Test_DataExchange_JsonImportHandlesLateColumnsAndHeaderCasing() {
            string filePath = Path.Combine(_directoryWithFiles, "DataExchange.JsonLateColumns.xlsx");
            const string json = "[{\"Name\":\"Alpha\"},{\"name\":\"Beta\",\"Score\":20}]";

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                string range = sheet.FromJson(json);
                Assert.Equal("A1:B3", range);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelSheet sheet = document.GetSheet("Data");
                DataTable table = sheet.ToDataTable("A1:B3");

                Assert.Equal(new[] { "Name", "Score" }, table.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray());
                Assert.Equal("Alpha", table.Rows[0]["Name"]);
                Assert.Equal(string.Empty, table.Rows[0]["Score"]);
                Assert.Equal("Beta", table.Rows[1]["Name"]);
                Assert.Equal(20D, Convert.ToDouble(table.Rows[1]["Score"]));
            }
        }

        [Fact]
        public void Test_DataExchange_JsonImportHandlesSparseLateWideColumns() {
            string filePath = Path.Combine(_directoryWithFiles, "DataExchange.JsonSparseLateWideColumns.xlsx");
            const string json = "[{\"Name\":\"Alpha\"},{\"Name\":\"Beta\",\"Metric25\":25},{\"Metric50\":50}]";

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                string range = sheet.FromJson(json);
                Assert.Equal("A1:C4", range);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelSheet sheet = document.GetSheet("Data");
                DataTable table = sheet.ToDataTable("A1:C4");

                Assert.Equal(new[] { "Name", "Metric25", "Metric50" }, table.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray());
                Assert.Equal("Alpha", table.Rows[0]["Name"]);
                Assert.Equal(string.Empty, table.Rows[0]["Metric25"]);
                Assert.Equal(string.Empty, table.Rows[0]["Metric50"]);
                Assert.Equal("Beta", table.Rows[1]["Name"]);
                Assert.Equal(25D, Convert.ToDouble(table.Rows[1]["Metric25"]));
                Assert.Equal(string.Empty, table.Rows[1]["Metric50"]);
                Assert.Equal(string.Empty, table.Rows[2]["Name"]);
                Assert.Equal(string.Empty, table.Rows[2]["Metric25"]);
                Assert.Equal(50D, Convert.ToDouble(table.Rows[2]["Metric50"]));
            }
        }

        [Fact]
        public void Test_DataExchange_JsonImportSparseRowsKeepLastDuplicateValue() {
            string filePath = Path.Combine(_directoryWithFiles, "DataExchange.JsonSparseDuplicateValues.xlsx");
            string columns = string.Join(",", Enumerable.Range(0, 40).Select(index => "\"Metric" + index + "\":" + index));
            string json = "[{" + columns + "},{\"Metric0\":\"Keep\",\"metric0\":null,\"Metric39\":39}]";

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                string range = sheet.FromJson(json);
                Assert.Equal("A1:AN3", range);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                ExcelSheet sheet = document.GetSheet("Data");
                DataTable table = sheet.ToDataTable("A1:AN3");

                Assert.Equal(string.Empty, table.Rows[1]["Metric0"]);
                Assert.Equal(39D, Convert.ToDouble(table.Rows[1]["Metric39"]));
            }
        }
    }
}

using System.Text;
using System.Data;
using System.Text.Json;
using OfficeIMO.Data.Conversions;
using Xunit;

namespace OfficeIMO.Data.Conversions.Tests;

// Tree <-> table contract: rows, path headers, canonical-only inference and the public DbDataReader bridge.
public class TableConversionTests {
    private const string Records = """
        [
          {"id": 1, "name": "Ann", "zip": "007", "price": 1.10, "big": 123456789012345678901234, "active": true,
           "address": {"city": "Gdansk", "street": null}, "tags": ["a", "b"], "born": "1990-05-01"},
          {"id": 2, "name": "Bob", "zip": "120", "price": 2, "big": 1, "active": false,
           "address": {"city": "Warsaw", "street": "Long"}, "tags": ["c"], "born": "1991-01-31"}
        ]
        """;

    [Fact]
    public void JsonRecordsSurviveCsvRoundTrip() {
        (string csv, _) = Convert.Run(Records, DataFormat.Json, DataFormat.Csv);
        Assert.StartsWith("id,name,zip,price,big,active,address.city,address.street,tags[0],tags[1],born\n", csv);
        (string json, _) = Convert.Run(csv, DataFormat.Csv, DataFormat.Json);
        using JsonDocument back = JsonDocument.Parse(json);
        JsonElement first = back.RootElement[0];
        Assert.Equal("007", first.GetProperty("zip").GetString());
        Assert.Equal("1.10", first.GetProperty("price").GetRawText());
        Assert.Equal("123456789012345678901234", first.GetProperty("big").GetRawText());
        Assert.True(first.GetProperty("active").GetBoolean());
        Assert.Equal("Gdansk", first.GetProperty("address").GetProperty("city").GetString());
        Assert.Equal(string.Empty, first.GetProperty("address").GetProperty("street").GetString());
        Assert.Equal(2, first.GetProperty("tags").GetArrayLength());
        Assert.Equal(1, back.RootElement[1].GetProperty("tags").GetArrayLength());
        Assert.Equal("1990-05-01", first.GetProperty("born").GetString());
    }

    // CSV output finds its columns from a sample; a later row with a new column fails instead of losing the value.
    [Fact]
    public void LateColumnsFailUnlessEveryRowIsRead() {
        const string rows = "[{\"a\":1},{\"a\":2},{\"a\":3,\"b\":4}]";
        var error = Assert.Throws<InvalidDataException>(() => Convert.Run(rows, DataFormat.Json, DataFormat.Csv, new DataConversionOptions { SampleRows = 2 }));
        Assert.Contains("'b'", error.Message);
        (string csv, _) = Convert.Run(rows, DataFormat.Json, DataFormat.Csv, new DataConversionOptions { SampleRows = 0 });
        Assert.Equal("a,b\n1,\n2,\n3,4\n", csv);
    }

    // Rows of a large JSON array cross the reader's buffer refills; each row is written once, in order, with its findings.
    [Fact]
    public void JsonRowsAcrossBufferRefillsAreWrittenOnce() {
        var json = new StringBuilder("[");
        var expected = new StringBuilder("id,name,note\n");
        for (int i = 0; i < 4000; i++) {
            if (i > 0) json.Append(',');
            json.Append("{\"id\":").Append(i).Append(",\"name\":\"caf\\u00e9 \\\"").Append(i).Append("\\\"\",\"note\":").Append(i % 1000 == 999 ? "\"42\"" : "null").Append('}');
            expected.Append(i).Append(",\"café \"\"").Append(i).Append("\"\"\",").Append(i % 1000 == 999 ? "42" : string.Empty).Append('\n');
        }

        (string csv, OfficeConversionReport report) = Convert.Run(json.Append(']').ToString(), DataFormat.Json, DataFormat.Csv);
        Assert.Equal(expected.ToString(), csv);
        Assert.Equal(4000, report.ItemsWritten);
        OfficeDiagnostic typed = Assert.Single(report.Diagnostics, d => d.Code == OfficeDiagnosticCodes.ScalarTypeDropped);
        Assert.Contains("(4 occurrences)", typed.Message);
    }

    [Fact]
    public void SelectReadsRowsBelowTheRootAndSeveralMembersAreTables() {
        const string response = "{\"data\":[{\"a\":1},{\"a\":2}],\"meta\":{\"page\":1}}";
        (string csv, _) = Convert.Run(response, DataFormat.Json, DataFormat.Csv, new DataConversionOptions { Select = "data" });
        Assert.Equal("a\n1\n2\n", csv);
        var error = Assert.Throws<InvalidDataException>(() => Convert.Run("{\"x\":[{\"a\":1}],\"y\":[{\"b\":2}]}", DataFormat.Json, DataFormat.Csv));
        Assert.Contains("x, y", error.Message);
    }

    [Fact]
    public void StructuredDataReaderFeedsAnyDataReaderConsumer() {
        using var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(Records));
        using StructuredDataReader reader = StructuredDataReader.Open(stream, OfficeDocumentFormat.Json);
        var table = new DataTable();
        table.Load(reader);
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(typeof(long), table.Columns["id"]!.DataType);
        Assert.Equal(typeof(decimal), table.Columns["price"]!.DataType);
        Assert.Equal("Warsaw", table.Rows[1]["address.city"]);
    }

    [Fact]
    public void MultiDocumentYamlBecomesRowsAndJsonLinesExplodesArrays() {
        (string csv, _) = Convert.Run("kind: Pod\nname: a\n---\nkind: Service\nname: b\n", DataFormat.Yaml, DataFormat.Csv);
        Assert.Equal("kind,name\nPod,a\nService,b\n", csv);
        (string lines, _) = Convert.Run("[{\"a\":1},{\"a\":2}]", DataFormat.Json, DataFormat.JsonLines);
        Assert.Equal("{\"a\":1}\n{\"a\":2}\n", lines);
        (string array, _) = Convert.Run(lines, DataFormat.JsonLines, DataFormat.Json);
        Assert.Equal(2, JsonDocument.Parse(array).RootElement.GetArrayLength());
    }
}

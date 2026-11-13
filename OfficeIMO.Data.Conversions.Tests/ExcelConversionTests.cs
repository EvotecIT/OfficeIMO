using System.Text.Json;
using System.Text.Json.Nodes;
using OfficeIMO.Data.Conversions;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Data.Conversions.Tests;

// Worksheets are the typed table target: rows round-trip through a workbook, and what a cell cannot hold is reported.
public class ExcelConversionTests {
    [Fact]
    public void JsonRowsRoundTripThroughAWorkbook() {
        JsonNode rows = JsonNode.Parse("""
            {"people":[{"id":1,"name":"Ann","score":1.5,"active":true,"address":{"city":"Gdansk"}},{"id":2,"name":"Bob","score":2,"active":false,"address":{"city":"Warsaw"}}],
             "teams":[{"team":"Red"}]}
            """)!;
        using var workbook = new MemoryStream();
        OfficeConversionReport written = rows.SaveAsExcel(workbook);
        Assert.False(written.HasLoss);
        workbook.Position = 0;
        using ExcelDocument document = ExcelDocument.Load(workbook);
        Assert.Equal(new[] { "people", "teams" }, document.Sheets.Select(s => s.Name).ToArray());
        JsonNode back = document.ToJson()!;
        Assert.Equal("Gdansk", back["people"]![0]!["address"]!["city"]!.GetValue<string>());
        Assert.Equal(1.5m, back["people"]![0]!["score"]!.GetValue<decimal>());
        Assert.False(back["people"]![1]!["active"]!.GetValue<bool>());
        Assert.Equal("Red", back["teams"]![0]!["team"]!.GetValue<string>());
    }

    [Fact]
    public void LongNumbersAndOffsetsAreReported() {
        JsonNode rows = JsonNode.Parse("[{\"id\":12345678901234567,\"at\":\"2026-11-10T10:00:00+01:00\"}]")!;
        using var workbook = new MemoryStream();
        OfficeConversionReport report = rows.SaveAsExcel(workbook);
        Assert.Contains(OfficeDiagnosticCodes.NumberPrecision, report.Diagnostics.Select(d => d.Code));
        workbook.Position = 0;
        using ExcelDocument document = ExcelDocument.Load(workbook);
        Assert.Equal("12345678901234567", document.ToJson()![0]!["id"]!.GetValue<string>());
    }
}

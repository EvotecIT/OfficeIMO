using System.Text.Json;
using System.Text.Json.Nodes;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class OfficeBatchReportTests {
    [Fact]
    public void CollectorRetainsBoundedItemsWithoutLosingAggregateCounts() {
        var collector = new OfficeBatchReportCollector();
        var item = Item("Small observation");
        for (int index = 0; index <= OfficeBatchReportCollector.MaximumItems; index++) collector.Report(item);
        var summary = new OfficeBatchSummary(10_001, 10_001, 0, 0, false);
        OfficeBatchReport report = collector.Complete(summary);
        Assert.True(report.Succeeded);
        Assert.Same(summary, report.Summary);
        Assert.Equal(10_000, report.Items.Count);
        Assert.True(report.ItemsTruncated);
        string json = OfficeBatchReportSerializer.Serialize(report);
        Assert.Equal(10_001, OfficeBatchReportSerializer.Deserialize(json).Summary.Completed);
    }

    [Fact]
    public void CollectorBoundsEncodedBytesAndRejectsOneOversizedObservation() {
        var collector = new OfficeBatchReportCollector();
        var oversized = Item(new string('x', OfficeBatchReportCollector.MaximumItemBytes + 1));
        collector.Report(oversized);
        Assert.Throws<InvalidDataException>(() => OfficeBatchReportSerializer.SerializeItemEvent(oversized));
        var encoded = Item(new string('\u0001', 400_000));
        for (int index = 0; index < 10; index++) collector.Report(encoded);
        OfficeBatchReport report = collector.Complete(new(11, 11, 0, 0, false));
        Assert.True(report.ItemsTruncated);
        Assert.InRange(report.Items.Count, 1, 9);
        Assert.Equal(11, report.Summary.Completed);
        string json = OfficeBatchReportSerializer.Serialize(report);
        Assert.True(System.Text.Encoding.UTF8.GetByteCount(json) <= OfficeBatchReportSerializer.MaximumReportBytes);
        Assert.Equal(report.Items.Count, OfficeBatchReportSerializer.Deserialize(json).Items.Count);
    }

    [Fact]
    public void BatchEventsAndReportRoundTripRetainDiagnosticsAndExplicitSuccess() {
        var collector = new OfficeBatchReportCollector();
        var item = Item("Preserved outcome") with {
            Diagnostics = [new OfficeDiagnostic(OfficeDiagnosticCodes.TableCellContentClipped, OfficeDiagnosticSeverity.Warning, "A table was simplified",
                "page 2", "conversion", OfficeConversionLossKind.Approximation)]
        };
        collector.Report(item);
        var summary = new OfficeBatchSummary(1, 1, 0, 0, false);
        OfficeBatchReport original = collector.Complete(summary);
        OfficeBatchReport restored = OfficeBatchReportSerializer.Deserialize(OfficeBatchReportSerializer.Serialize(original));
        Assert.True(restored.Succeeded);
        Assert.False(restored.ItemsTruncated);
        var diagnostic = Assert.Single(Assert.Single(restored.Items).Diagnostics);
        Assert.Equal("page 2", diagnostic.Location);
        Assert.Equal(OfficeDiagnosticSeverity.Warning, diagnostic.Severity);
        using var itemEvent = JsonDocument.Parse(OfficeBatchReportSerializer.SerializeItemEvent(item));
        Assert.Equal("item", itemEvent.RootElement.GetProperty("type").GetString());
        Assert.Equal("Completed", itemEvent.RootElement.GetProperty("item").GetProperty("status").GetString());
        using var summaryEvent = JsonDocument.Parse(OfficeBatchReportSerializer.SerializeSummaryEvent(summary, succeeded: false));
        Assert.Equal("summary", summaryEvent.RootElement.GetProperty("type").GetString());
        Assert.False(summaryEvent.RootElement.GetProperty("succeeded").GetBoolean());
        Assert.Equal(1, summaryEvent.RootElement.GetProperty("summary").GetProperty("completed").GetInt64());
        Assert.False(collector.Complete(new(0, 0, 0, 0, false)).Succeeded);
        Assert.True(collector.Complete(new(0, 0, 0, 0, false), allowEmpty: true).Succeeded);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void StrictReportReaderRejectsUnknownDuplicateNumericAndOversizedInput(int invalidCase) {
        var collector = new OfficeBatchReportCollector(); collector.Report(Item("Observation"));
        string json = OfficeBatchReportSerializer.Serialize(collector.Complete(new(1, 1, 0, 0, false)));
        if (invalidCase == 0) json = json.Insert(1, "\"unknown\":true,");
        else if (invalidCase == 1) json = json.Insert(1, "\"schema\":\"officeimo.conversion.batch.v1\",");
        else if (invalidCase == 2) json = json.Replace("\"Completed\"", "0");
        else {
            Assert.Throws<InvalidDataException>(() => OfficeBatchReportSerializer.Deserialize(new string(' ', OfficeBatchReportSerializer.MaximumReportBytes + 1)));
            return;
        }
        Assert.Throws<JsonException>(() => OfficeBatchReportSerializer.Deserialize(json));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void MalformedDiagnosticConstructorValuesAreRejectedAsInvalidReportData(int invalidCase) {
        var collector = new OfficeBatchReportCollector();
        collector.Report(Item("Observed outcome") with { Diagnostics = [new OfficeDiagnostic(OfficeDiagnosticCodes.TableCellContentClipped, OfficeDiagnosticSeverity.Warning, "A layout warning")] });
        string json = OfficeBatchReportSerializer.Serialize(collector.Complete(new(1, 1, 0, 0, false)));
        JsonNode data = JsonNode.Parse(json)!;
        JsonObject diagnostic = data["items"]![0]!["diagnostics"]![0]!.AsObject();
        diagnostic[invalidCase % 2 == 0 ? "code" : "message"] = invalidCase < 2 ? null : " ";
        var error = Assert.Throws<InvalidDataException>(() => OfficeBatchReportSerializer.Deserialize(data.ToJsonString()));
        Assert.IsAssignableFrom<ArgumentException>(error.InnerException);
    }

    [Fact]
    public void ReportReadersPreserveNullCallerArgumentErrors() {
        Assert.Throws<ArgumentNullException>(() => OfficeBatchReportSerializer.Deserialize(null!));
        Assert.Throws<ArgumentNullException>(() => OfficeBatchReport.Parse(null!));
    }

    private static OfficeBatchItemReport Item(string message) => new("source.txt", "output.pdf", OfficeBatchItemStatus.Completed, false, message, []);
}

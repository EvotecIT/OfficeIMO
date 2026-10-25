using System.Text.Json;
using System.Text.Json.Nodes;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class OfficeConversionBatchReportTests {
    [Fact]
    public void CollectorRetainsBoundedItemsWithoutLosingAggregateCounts() {
        var collector = new OfficeConversionBatchReportCollector();
        var item = Item("Small observation");
        for (int index = 0; index <= OfficeConversionBatchReportCollector.MaximumItems; index++) collector.Report(item);
        var summary = new OfficeConversionBatchResult(10_001, 10_001, 0, 0, false);
        OfficeConversionBatchReport report = collector.Complete(summary);
        Assert.True(report.Succeeded);
        Assert.Same(summary, report.Summary);
        Assert.Equal(10_000, report.Items.Count);
        Assert.True(report.ItemsTruncated);
        string json = OfficeConversionBatchReportSerializer.Serialize(report);
        Assert.Equal(10_001, OfficeConversionBatchReportSerializer.Deserialize(json).Summary.Completed);
    }

    [Fact]
    public void CollectorBoundsEncodedBytesAndRejectsOneOversizedObservation() {
        var collector = new OfficeConversionBatchReportCollector();
        var oversized = Item(new string('x', OfficeConversionBatchReportCollector.MaximumItemBytes + 1));
        collector.Report(oversized);
        Assert.Throws<InvalidDataException>(() => OfficeConversionBatchReportSerializer.SerializeItemEvent(oversized));
        var encoded = Item(new string('\u0001', 400_000));
        for (int index = 0; index < 10; index++) collector.Report(encoded);
        OfficeConversionBatchReport report = collector.Complete(new(11, 11, 0, 0, false));
        Assert.True(report.ItemsTruncated);
        Assert.InRange(report.Items.Count, 1, 9);
        Assert.Equal(11, report.Summary.Completed);
        string json = OfficeConversionBatchReportSerializer.Serialize(report);
        Assert.True(System.Text.Encoding.UTF8.GetByteCount(json) <= OfficeConversionBatchReportSerializer.MaximumReportBytes);
        Assert.Equal(report.Items.Count, OfficeConversionBatchReportSerializer.Deserialize(json).Items.Count);
    }

    [Fact]
    public void BatchEventsAndReportRoundTripRetainDiagnosticsAndExplicitSuccess() {
        var collector = new OfficeConversionBatchReportCollector();
        var item = Item("Preserved outcome") with {
            Diagnostics = [new OfficeWorkflowDiagnostic("LayoutChanged", "A table was simplified", OfficeWorkflowDiagnosticSeverity.Warning,
                "conversion", new Dictionary<string, string> { ["location"] = "page 2", ["lossKind"] = "Approximation" })]
        };
        collector.Report(item);
        var summary = new OfficeConversionBatchResult(1, 1, 0, 0, false);
        OfficeConversionBatchReport original = collector.Complete(summary);
        OfficeConversionBatchReport restored = OfficeConversionBatchReportSerializer.Deserialize(OfficeConversionBatchReportSerializer.Serialize(original));
        Assert.True(restored.Succeeded);
        Assert.False(restored.ItemsTruncated);
        var diagnostic = Assert.Single(Assert.Single(restored.Items).Diagnostics);
        Assert.Equal("page 2", diagnostic.Details["location"]);
        Assert.Equal(OfficeWorkflowDiagnosticSeverity.Warning, diagnostic.Severity);
        using var itemEvent = JsonDocument.Parse(OfficeConversionBatchReportSerializer.SerializeItemEvent(item));
        Assert.Equal("item", itemEvent.RootElement.GetProperty("type").GetString());
        Assert.Equal("Completed", itemEvent.RootElement.GetProperty("item").GetProperty("status").GetString());
        using var summaryEvent = JsonDocument.Parse(OfficeConversionBatchReportSerializer.SerializeSummaryEvent(summary, succeeded: false));
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
        var collector = new OfficeConversionBatchReportCollector(); collector.Report(Item("Observation"));
        string json = OfficeConversionBatchReportSerializer.Serialize(collector.Complete(new(1, 1, 0, 0, false)));
        if (invalidCase == 0) json = json.Insert(1, "\"unknown\":true,");
        else if (invalidCase == 1) json = json.Insert(1, "\"schema\":\"officeimo.conversion.batch.v1\",");
        else if (invalidCase == 2) json = json.Replace("\"Completed\"", "0");
        else {
            Assert.Throws<InvalidDataException>(() => OfficeConversionBatchReportSerializer.Deserialize(new string(' ', OfficeConversionBatchReportSerializer.MaximumReportBytes + 1)));
            return;
        }
        Assert.Throws<JsonException>(() => OfficeConversionBatchReportSerializer.Deserialize(json));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void MalformedDiagnosticConstructorValuesAreRejectedAsInvalidReportData(int invalidCase) {
        var collector = new OfficeConversionBatchReportCollector();
        collector.Report(Item("Observed outcome") with { Diagnostics = [new OfficeWorkflowDiagnostic("LayoutChanged", "A layout warning")] });
        string json = OfficeConversionBatchReportSerializer.Serialize(collector.Complete(new(1, 1, 0, 0, false)));
        JsonNode data = JsonNode.Parse(json)!;
        JsonObject diagnostic = data["items"]![0]!["diagnostics"]![0]!.AsObject();
        diagnostic[invalidCase % 2 == 0 ? "code" : "message"] = invalidCase < 2 ? null : " ";
        var error = Assert.Throws<InvalidDataException>(() => OfficeConversionBatchReportSerializer.Deserialize(data.ToJsonString()));
        Assert.IsAssignableFrom<ArgumentException>(error.InnerException);
    }

    [Fact]
    public void ReportReadersPreserveNullCallerArgumentErrors() {
        Assert.Throws<ArgumentNullException>(() => OfficeConversionBatchReportSerializer.Deserialize(null!));
        Assert.Throws<ArgumentNullException>(() => OfficeConversionJobRunReportSerializer.Deserialize(null!));
    }

    private static OfficeConversionBatchItemResult Item(string message) => new("source.txt", "output.pdf", OfficeWorkflowStatus.Completed, false, message, []);
}

using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Workflow;

/// <summary>Serializes concurrent batch outcomes without mixing human progress into stdout.</summary>
internal sealed class WorkflowBatchOutput(TextWriter output, TextWriter error, string? format) : IProgress<OfficeConversionBatchItemResult> {
    private const int MaximumReportItems = 10_000;
    private const long MaximumReportItemBytes = 16L * 1024 * 1024;
    private const string Schema = "officeimo.conversion.batch.v1";
    private readonly object _gate = new();
    private readonly List<OfficeConversionBatchItemResult> _items = [];
    private long _reported;
    private long _retainedBytes;

    public void Report(OfficeConversionBatchItemResult item) {
        lock (_gate) {
            _reported++;
            if (format == "ndjson") output.WriteLine(JsonSerializer.Serialize(
                new WorkflowBatchItemEvent(Schema, "item", item), WorkflowBatchOutputJsonContext.Default.WorkflowBatchItemEvent));
            else if (format == "json" && _items.Count < MaximumReportItems) {
                int bytes = JsonSerializer.SerializeToUtf8Bytes(item, WorkflowBatchOutputJsonContext.Default.OfficeConversionBatchItemResult).Length;
                if (bytes <= MaximumReportItemBytes - _retainedBytes) { _items.Add(item); _retainedBytes += bytes; }
            }
            if (item.Skipped || item.Status != OfficeWorkflowStatus.Completed) error.WriteLine(item.InputPath + ": " + item.Summary);
            foreach (var finding in item.Diagnostics.Where(finding => finding.Severity != OfficeWorkflowDiagnosticSeverity.Information)) {
                finding.Details.TryGetValue("location", out string? location);
                error.WriteLine(item.InputPath + ": " + finding.Code + (location == null ? "" : " [" + location + "]") + ": " + finding.Message);
            }
        }
    }

    internal void Complete(OfficeConversionBatchResult result, bool allowEmpty = false) {
        lock (_gate) {
            bool succeeded = !result.Cancelled && result.FailureSummary == null && result.Failed == 0 && (allowEmpty || result.Selected > 0);
            if (result.FailureSummary != null) error.WriteLine("Batch stopped: " + result.FailureSummary);
            if (result.Selected == 0 && !allowEmpty && !result.Cancelled && result.FailureSummary == null)
                error.WriteLine("No files were selected for conversion. Check the inputs and executable route, or use --allow-empty for an intentionally empty job.");
            if (format == "json") {
                output.WriteLine(JsonSerializer.Serialize(new WorkflowBatchReport(Schema, "report", succeeded, result,
                    _items.ToArray(), _reported > _items.Count), WorkflowBatchOutputJsonContext.Default.WorkflowBatchReport));
            } else if (format == "ndjson") {
                output.WriteLine(JsonSerializer.Serialize(new WorkflowBatchSummaryEvent(Schema, "summary", succeeded, result),
                    WorkflowBatchOutputJsonContext.Default.WorkflowBatchSummaryEvent));
            } else output.WriteLine(OfficeConversionBatchSerializer.SerializeResult(result));
        }
    }
}

internal sealed record WorkflowBatchItemEvent(string Schema, string Type, OfficeConversionBatchItemResult Item);
internal sealed record WorkflowBatchSummaryEvent(string Schema, string Type, bool Succeeded, OfficeConversionBatchResult Summary);
internal sealed record WorkflowBatchReport(string Schema, string Type, bool Succeeded, OfficeConversionBatchResult Summary,
    OfficeConversionBatchItemResult[] Items, bool ItemsTruncated);

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase, UseStringEnumConverter = true,
    GenerationMode = JsonSourceGenerationMode.Serialization)]
[JsonSerializable(typeof(WorkflowBatchItemEvent))]
[JsonSerializable(typeof(WorkflowBatchSummaryEvent))]
[JsonSerializable(typeof(WorkflowBatchReport))]
[JsonSerializable(typeof(OfficeConversionBatchItemResult))]
internal partial class WorkflowBatchOutputJsonContext : JsonSerializerContext { }

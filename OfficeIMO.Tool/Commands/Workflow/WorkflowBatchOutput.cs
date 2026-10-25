using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Workflow;

/// <summary>Serializes concurrent batch outcomes without mixing human progress into stdout.</summary>
internal sealed class WorkflowBatchOutput(TextWriter output, TextWriter error, string? format) : IProgress<OfficeConversionBatchItemResult> {
    private readonly object _gate = new();
    private readonly OfficeConversionBatchReportCollector _report = new();

    public void Report(OfficeConversionBatchItemResult item) {
        lock (_gate) {
            if (format == "ndjson") output.WriteLine(OfficeConversionBatchReportSerializer.SerializeItemEvent(item));
            else if (format == "json") _report.Report(item);
            if (item.Skipped || item.Status != OfficeWorkflowStatus.Completed) error.WriteLine(item.InputPath + ": " + item.Summary);
            foreach (var finding in item.Diagnostics.Where(finding => finding.Severity != OfficeWorkflowDiagnosticSeverity.Information)) {
                finding.Details.TryGetValue("location", out string? location);
                error.WriteLine(item.InputPath + ": " + finding.Code + (location == null ? "" : " [" + location + "]") + ": " + finding.Message);
            }
        }
    }

    internal void Complete(OfficeConversionBatchResult result, bool allowEmpty = false, string? recordingError = null) {
        lock (_gate) {
            var report = _report.Complete(result, allowEmpty);
            bool succeeded = report.Succeeded && recordingError == null;
            if (recordingError != null) error.WriteLine("The conversions finished, but their run report could not be finalized: " + recordingError);
            if (result.FailureSummary != null) error.WriteLine("Batch stopped: " + result.FailureSummary);
            if (result.Selected == 0 && !allowEmpty && !result.Cancelled && result.FailureSummary == null)
                error.WriteLine("No files were selected for conversion. Check the inputs and executable route, or use --allow-empty for an intentionally empty job.");
            if (format == "json") {
                output.WriteLine(OfficeConversionBatchReportSerializer.Serialize(report with { Succeeded = succeeded }));
            } else if (format == "ndjson") {
                output.WriteLine(OfficeConversionBatchReportSerializer.SerializeSummaryEvent(result, succeeded));
            } else output.WriteLine(OfficeConversionBatchSerializer.SerializeResult(result));
        }
    }
}
